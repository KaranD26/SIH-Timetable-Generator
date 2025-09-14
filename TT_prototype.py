# Prototype Timetable Generator (GUI + OR-Tools)
# Requirements: pip install ortools
# Run: python TT_prototype.py

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from ortools.sat.python import cp_model
import csv
import threading
import queue
from typing import List, Tuple, Dict, Any

DEFAULT_SUBJECTS = [
    ("Math", 3, 2),
    ("Physics", 4, 2),  # +1 lab included
    ("Computer Programming", 4, 2),
    ("AI/ML", 4, 2),
    ("Computer Organisation", 3, 2),
    ("Discrete Math", 3, 2),
    ("Ethics", 3, 2),
]

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"]
SLOTS = ["9-10", "10-11", "11-12", "12-1", "2-3", "3-4", "4-5", "5-6"]  # Lunch 1-2 excluded
DEFAULT_NUM_BATCHES = 5
DEFAULT_NUM_ROOMS = 3
DEFAULT_MAX_PER_DAY = 6
DEFAULT_MIN_PER_DAY = 3
DEFAULT_MAX_PER_WEEK = 24
SOLVER_TIME_LIMIT = 40

# --- Solver Objective Weights ---
# These weights guide the solver towards a "good" timetable.
# Higher values mean the solver will prioritize that goal more.
WEIGHT_FULL_SLOT = 1000  # Prioritize using all rooms in a slot.
WEIGHT_ADJACENCY = 5     # Encourage back-to-back classes to minimize gaps.
WEIGHT_PENALTY_SAME_DAY_SUBJECT = 10 # Penalize scheduling the same subject multiple times on the same day.
WEIGHT_PENALTY_FREE_DAY = 2000 # Strongly penalize having a day with zero classes (encourages a compact week).

def flatten_subjects_from_listbox(tree):
    """Returns list of tuples (name, hours, faculty) from a Treeview."""
    items = []
    for iid in tree.get_children():
        name = tree.item(iid, "values")[0]
        hours = int(tree.item(iid, "values")[1])
        faculty = int(tree.item(iid, "values")[2])
        items.append((name, hours, faculty))
    return items

def solve_timetable(
    subjects: List[Tuple[str, int, int]],
    num_batches: int,
    num_rooms: int,
    max_per_day: int,
    max_per_week: int,
    min_per_day: int,
    days: List[str] = DAYS,
    slots: List[str] = SLOTS,
    time_limit: int = SOLVER_TIME_LIMIT,
) -> Tuple[Dict[int, List[str]], Dict[int, List[Tuple[int, int]]]]:
    """
    Generates a timetable using the CP-SAT solver.
    Returns: A tuple containing the schedule and room assignments.
    """
    # Basic feasibility checks
    num_subjects = len(subjects)
    num_days = len(days)
    num_slots_per_day = len(slots)
    total_slots = num_days * num_slots_per_day

    if max_per_day > num_slots_per_day:
        raise ValueError(f"Max hours/day ({max_per_day}) exceeds available slots per day ({num_slots_per_day}).")
    if max_per_week > total_slots:
        raise ValueError(f"Max hours/week ({max_per_week}) exceeds total slots in the week ({total_slots}).")
    if min_per_day > max_per_day:
        raise ValueError(f"Min hours/day ({min_per_day}) cannot be greater than Max hours/day ({max_per_day}).")

    required_per_batch = sum(hours for (_, hours, _) in subjects)
    if required_per_batch > max_per_week:
        raise ValueError(f"Total required hours per batch ({required_per_batch}) exceeds the weekly limit ({max_per_week}).")

    total_required_all_batches = required_per_batch * num_batches
    max_possible_with_rooms = num_rooms * total_slots
    if total_required_all_batches > max_possible_with_rooms:
        raise ValueError(f"Not enough rooms. Required class-slots ({total_required_all_batches}) "
                         f"exceed total room capacity ({max_possible_with_rooms}).")

    model = cp_model.CpModel()

    # --- Decision Variables ---
    x = {}  # x[b, s, t] = 1 if batch b has subject s at time t
    for b in range(num_batches):
        for s_idx in range(num_subjects):
            for t in range(total_slots):
                x[(b, s_idx, t)] = model.NewBoolVar(f"x_b{b}_s{s_idx}_t{t}")

    # y[b, t] is a helper variable, true if batch b has ANY class at time t.
    y = {}
    for b in range(num_batches):
        for t in range(total_slots):
            y[(b, t)] = model.NewBoolVar(f"y_b{b}_t{t}")
            # A batch can only have one subject at a time. This also links y to x.
            model.Add(sum(x[(b, s_idx, t)] for s_idx in range(num_subjects)) == y[(b, t)])

    # --- Hard Constraints ---
    # Per-batch per-day limit and enforce min_per_day if day is active
    free_day_penalty_vars = []
    for b in range(num_batches):
        for d in range(num_days):
            day_ts = [d * num_slots_per_day + k for k in range(num_slots_per_day)]
            daily_hours = sum(y[(b, t)] for t in day_ts)
            model.Add(daily_hours <= max_per_day)

            # To enforce min_per_day only on days that are not free, we need an indicator.
            is_active = model.NewBoolVar(f"day_active_b{b}_d{d}")
            # is_active is true if there's at least one class that day.
            model.Add(daily_hours >= 1).OnlyEnforceIf(is_active)
            model.Add(daily_hours == 0).OnlyEnforceIf(is_active.Not())

            # This variable is used in the objective to penalize having free days.
            free_day_penalty_vars.append(is_active.Not())

            if min_per_day > 0:
                model.Add(daily_hours >= min_per_day).OnlyEnforceIf(is_active)

    # Per-batch per-week limit
    for b in range(num_batches):
        model.Add(sum(y[(b, t)] for t in range(total_slots)) <= max_per_week)

    # Subject weekly requirements per batch (exact)
    for b in range(num_batches):
        for s_idx, (_, hours, _) in enumerate(subjects):
            model.Add(sum(x[(b, s_idx, t)] for t in range(total_slots)) == hours)

    # Room capacity per slot
    for t in range(total_slots):
        model.Add(sum(y[(b, t)] for b in range(num_batches)) <= num_rooms)

    # Faculty per subject constraint
    for s_idx, (_, _, faculty_count) in enumerate(subjects):
        for t in range(total_slots):
            model.Add(sum(x[(b, s_idx, t)] for b in range(num_batches)) <= faculty_count)

    # --- Soft Constraints (for Objective Function) ---
    # Soft constraint: avoid scheduling the same subject twice on the same day for a batch.
    same_day_penalty_terms = []
    for b in range(num_batches):
        for s_idx, (_, hours, _) in enumerate(subjects):
            # Only penalize if a subject needs to be taught on more than one day.
            if hours <= num_days:
                for d in range(num_days):
                    daily_subject_hours = sum(x[(b, s_idx, d * num_slots_per_day + k)] for k in range(num_slots_per_day))
                    # Penalize if daily_subject_hours > 1.
                    penalty_var = model.NewIntVar(0, num_slots_per_day - 1, f"sameday_penalty_b{b}_s{s_idx}_d{d}")
                    model.Add(penalty_var >= daily_subject_hours - 1)
                    same_day_penalty_terms.append(penalty_var)

    # Maximize fully-occupied slots (i.e., use rooms efficiently)
    slot_filled = {}
    for t in range(total_slots):
        slot_filled[t] = model.NewIntVar(0, num_rooms, f"slot_filled_{t}")
        model.Add(slot_filled[t] == sum(y[(b, t)] for b in range(num_batches)))

    # full_slot[t] is true if all rooms are occupied at time t.
    full_slot = {}
    for t in range(total_slots):
        full_slot[t] = model.NewBoolVar(f"full_slot_{t}")
        # This is a standard way to implement an indicator: full_slot[t] <=> (slot_filled[t] == num_rooms)
        model.Add(slot_filled[t] == num_rooms).OnlyEnforceIf(full_slot[t])
        model.Add(slot_filled[t] != num_rooms).OnlyEnforceIf(full_slot[t].Not())

    # Adjacency variables to encourage contiguous classes (reduce gaps)
    adj_list = []
    for b in range(num_batches):
        for d in range(num_days):
            for k in range(num_slots_per_day - 1):
                t = d * num_slots_per_day + k
                adj_var = model.NewBoolVar(f"adj_b{b}_d{d}_k{k}")
                # adj_var is true if there's a class at time t AND t+1
                model.AddBoolAnd([y[(b, t)], y[(b, t + 1)]]).OnlyEnforceIf(adj_var)
                model.AddBoolOr([y[(b, t)].Not(), y[(b, t + 1)].Not()]).OnlyEnforceIf(adj_var.Not())
                adj_list.append(adj_var)

    # --- Objective Function ---
    # Maximize full slots (primary), then adjacency; penalize same-day repetition & free days
    model.Maximize(
        sum(WEIGHT_FULL_SLOT * v for v in full_slot.values()) +
        sum(WEIGHT_ADJACENCY * v for v in adj_list) -
        sum(WEIGHT_PENALTY_SAME_DAY_SUBJECT * term for term in same_day_penalty_terms) -
        sum(WEIGHT_PENALTY_FREE_DAY * v for v in free_day_penalty_vars)
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers = 8

    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("Solver found no feasible solution within time limit.")

    # Build schedule per batch
    schedule = {}
    for b in range(num_batches):
        schedule[b] = ["Free"] * total_slots
        for s_idx, (name, _, _) in enumerate(subjects):
            for t in range(total_slots):
                if solver.Value(x[(b, s_idx, t)]) == 1:
                    schedule[b][t] = name

    # Room assignments: for each slot, assign rooms 1..num_rooms to occupied batches in arbitrary order
    room_assignments = {}
    for t in range(total_slots):
        occupied_batches = [b for b in range(num_batches) if solver.Value(y[(b, t)]) == 1]
        assigns = []
        for idx, b in enumerate(occupied_batches):
            assigns.append((idx + 1, b))
        room_assignments[t] = assigns

    return schedule, room_assignments

class TimetableApp:
    def __init__(self, root):
        self.root = root
        root.title("Timetable Prototype - GUI + OR-Tools (Finished)")
        root.geometry("1200x750")

        param_frame = ttk.LabelFrame(root, text="Parameters")
        param_frame.pack(fill="x", padx=8, pady=6)

        ttk.Label(param_frame, text="Number of Batches:").grid(row=0, column=0, padx=6, pady=4, sticky="w")
        self.num_batches_var = tk.IntVar(value=DEFAULT_NUM_BATCHES)
        ttk.Spinbox(param_frame, from_=1, to=100, textvariable=self.num_batches_var, width=5).grid(row=0, column=1, sticky="w")

        ttk.Label(param_frame, text="Number of Rooms:").grid(row=0, column=2, padx=6, pady=4, sticky="w")
        self.num_rooms_var = tk.IntVar(value=DEFAULT_NUM_ROOMS)
        ttk.Spinbox(param_frame, from_=1, to=100, textvariable=self.num_rooms_var, width=5).grid(row=0, column=3, sticky="w")

        ttk.Label(param_frame, text="Max hours/day per batch:").grid(row=0, column=4, padx=6, pady=4, sticky="w")
        self.max_per_day_var = tk.IntVar(value=DEFAULT_MAX_PER_DAY)
        ttk.Spinbox(param_frame, from_=1, to=len(SLOTS), textvariable=self.max_per_day_var, width=5).grid(row=0, column=5, sticky="w")

        ttk.Label(param_frame, text="Min hours/day per batch:").grid(row=0, column=6, padx=6, pady=4, sticky="w")
        self.min_per_day_var = tk.IntVar(value=DEFAULT_MIN_PER_DAY)
        ttk.Spinbox(param_frame, from_=0, to=len(SLOTS), textvariable=self.min_per_day_var, width=5).grid(row=0, column=7, sticky="w")

        ttk.Label(param_frame, text="Max hours/week per batch:").grid(row=0, column=8, padx=6, pady=4, sticky="w")
        self.max_per_week_var = tk.IntVar(value=DEFAULT_MAX_PER_WEEK)
        ttk.Spinbox(param_frame, from_=1, to=len(DAYS)*len(SLOTS), textvariable=self.max_per_week_var, width=5).grid(row=0, column=9, sticky="w")

        ttk.Label(param_frame, text="Solver time limit (s):").grid(row=0, column=10, padx=6, pady=4, sticky="w")
        self.time_limit_var = tk.IntVar(value=SOLVER_TIME_LIMIT)
        ttk.Spinbox(param_frame, from_=5, to=300, textvariable=self.time_limit_var, width=5).grid(row=0, column=11, sticky="w")

        mid_frame = ttk.Frame(root)
        mid_frame.pack(fill="x", padx=8, pady=4)

        subj_frame = ttk.LabelFrame(mid_frame, text="Subjects (weekly hours & faculty count)")
        subj_frame.pack(side="left", fill="both", expand=False, padx=6, pady=4)

        columns = ("Subject", "Hours", "Faculty")
        self.subj_tree = ttk.Treeview(subj_frame, columns=columns, show="headings", height=10)
        for col in columns:
            self.subj_tree.heading(col, text=col)
            self.subj_tree.column(col, width=140 if col == "Subject" else 60, anchor="center")
        self.subj_tree.pack(side="left", padx=6, pady=6)
        self.subj_tree.bind("<<TreeviewSelect>>", self.on_subject_select)

        for name, hrs, faculty in DEFAULT_SUBJECTS:
            self.subj_tree.insert("", "end", values=(name, hrs, faculty))

        ctrl_frame = ttk.Frame(subj_frame)
        ctrl_frame.pack(side="left", fill="y", padx=6)

        ttk.Label(ctrl_frame, text="Name:").pack(anchor="w")
        self.subj_name_var = tk.StringVar()
        ttk.Entry(ctrl_frame, textvariable=self.subj_name_var, width=24).pack(anchor="w", pady=2)

        ttk.Label(ctrl_frame, text="Hours/week:").pack(anchor="w")
        self.subj_hours_var = tk.IntVar(value=3)
        ttk.Entry(ctrl_frame, textvariable=self.subj_hours_var, width=6).pack(anchor="w", pady=2)

        ttk.Label(ctrl_frame, text="Faculty Count:").pack(anchor="w")
        self.subj_faculty_var = tk.IntVar(value=2)
        ttk.Entry(ctrl_frame, textvariable=self.subj_faculty_var, width=6).pack(anchor="w", pady=2)

        ttk.Button(ctrl_frame, text="Add Subject", command=self.add_subject).pack(fill="x", pady=4)
        ttk.Button(ctrl_frame, text="Edit Selected", command=self.edit_subject).pack(fill="x", pady=2)
        ttk.Button(ctrl_frame, text="Remove Selected", command=self.remove_subject).pack(fill="x", pady=2)
        ttk.Button(ctrl_frame, text="Reset Defaults", command=self.reset_defaults).pack(fill="x", pady=6)

        action_frame = ttk.LabelFrame(mid_frame, text="Actions")
        action_frame.pack(side="left", fill="both", expand=True, padx=6, pady=4)

        ttk.Button(action_frame, text="Generate Timetable", command=self.generate_timetable, style="Accent.TButton").pack(padx=8, pady=12)
        ttk.Button(action_frame, text="Clear Output", command=self.clear_output).pack(padx=8, pady=6)
        ttk.Button(action_frame, text="Export to CSV", command=self.export_csv).pack(padx=8, pady=6)

        out_frame = ttk.LabelFrame(root, text="Output")
        out_frame.pack(fill="both", expand=True, padx=8, pady=6)

        self.text_area = scrolledtext.ScrolledText(out_frame, wrap=tk.NONE, font=("Consolas", 10))
        self.text_area.pack(fill="both", expand=True)

        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x")

        self.last_schedule_data = None
        self.solver_queue = queue.Queue()

    def add_subject(self):
        name = self.subj_name_var.get().strip()
        try:
            hours = int(self.subj_hours_var.get())
            faculty = int(self.subj_faculty_var.get())
        except ValueError:
            messagebox.showerror("Invalid Number", "Hours and Faculty Count must be integers.")
            return
        if not name:
            messagebox.showerror("Invalid name", "Subject name cannot be empty.")
            return
        if hours <= 0 or faculty <= 0:
            messagebox.showerror("Invalid Number", "Hours and Faculty Count must be positive.")
            return
        self.subj_tree.insert("", "end", values=(name, hours, faculty))
        self.subj_name_var.set("")
        self.subj_hours_var.set(3)
        self.subj_faculty_var.set(2)
        self.subj_tree.selection_remove(self.subj_tree.selection())

    def edit_subject(self):
        sel = self.subj_tree.selection()
        if not sel:
            messagebox.showerror("Select", "Select a subject to edit.")
            return
        iid = sel[0]
        name = self.subj_name_var.get().strip()
        try:
            hours = int(self.subj_hours_var.get())
            faculty = int(self.subj_faculty_var.get())
        except ValueError:
            messagebox.showerror("Invalid Number", "Hours and Faculty Count must be integers.")
            return
        if not name:
            messagebox.showerror("Invalid name", "Subject name cannot be empty.")
            return
        if hours <= 0 or faculty <= 0:
            messagebox.showerror("Invalid Number", "Hours and Faculty Count must be positive.")
            return
        self.subj_tree.item(iid, values=(name, hours, faculty))
        self.subj_name_var.set("")
        self.subj_hours_var.set(3)
        self.subj_faculty_var.set(2)
        self.subj_tree.selection_remove(sel)

    def on_subject_select(self, event=None):
        """When a user selects a subject, populate the entry fields for easy editing."""
        selection = self.subj_tree.selection()
        if not selection:
            return
        item = self.subj_tree.item(selection[0], "values")
        self.subj_name_var.set(item[0])
        self.subj_hours_var.set(int(item[1]))
        self.subj_faculty_var.set(int(item[2]))

    def remove_subject(self):
        sel = self.subj_tree.selection()
        if not sel:
            messagebox.showerror("Select", "Select a subject to remove.")
            return
        for iid in sel:
            self.subj_tree.delete(iid)

    def reset_defaults(self):
        for iid in self.subj_tree.get_children():
            self.subj_tree.delete(iid)
        for name, hrs, faculty in DEFAULT_SUBJECTS:
            self.subj_tree.insert("", "end", values=(name, hrs, faculty))

    def clear_output(self):
        self.text_area.delete(1.0, tk.END)
        self.status_var.set("Cleared output.")

    def export_csv(self):
        if not self.last_schedule_data:
            messagebox.showerror("No output", "No timetable to export. Generate first.")
            return

        schedule = self.last_schedule_data["schedule"]
        batch_names = self.last_schedule_data["batch_names"]
        days = self.last_schedule_data["days"]
        slots = self.last_schedule_data["slots"]
        num_days, num_slots_per_day, total_slots = len(days), len(slots), len(days) * len(slots)

        try:
            with open("timetable_output.csv", "w", newline="", encoding="utf8") as f:
                writer = csv.writer(f)
                header = ["Batch/Slot"] + [f"{days[d]} {slots[k]}" for d in range(num_days) for k in range(num_slots_per_day)]
                writer.writerow(header)
                for b, batch_name in enumerate(batch_names):
                    row = [batch_name]
                    for t in range(total_slots):
                        cell = schedule.get(b, ["Free"] * total_slots)[t]
                        row.append(cell)
                    writer.writerow(row)
            messagebox.showinfo("Exported", "Saved timetable_output.csv in current folder.")
        except Exception as e:
            messagebox.showerror("Export error", str(e))

    def _get_and_validate_params(self):
        """Reads parameters from the GUI, validates them, and returns them or None on error."""
        try:
            num_batches = int(self.num_batches_var.get())
            num_rooms = int(self.num_rooms_var.get())
            max_per_day = int(self.max_per_day_var.get())
            min_per_day = int(self.min_per_day_var.get())
            max_per_week = int(self.max_per_week_var.get())
            time_limit = int(self.time_limit_var.get())
        except ValueError:
            messagebox.showerror("Invalid params", "Numeric parameters must be integers.")
            return None

        if num_batches <= 0 or num_rooms <= 0 or max_per_day <= 0 or max_per_week <= 0 or min_per_day < 0:
            messagebox.showerror("Invalid params", "Numeric parameters must be positive (Min hours/day can be 0).")
            return None

        if min_per_day > max_per_day:
            messagebox.showerror("Invalid params", "Min hours/day cannot be greater than Max hours/day.")
            return None

        subjects = flatten_subjects_from_listbox(self.subj_tree)
        if not subjects:
            messagebox.showerror("No subjects", "Add at least one subject.")
            return None

        return {
            "subjects": subjects, "num_batches": num_batches, "num_rooms": num_rooms,
            "max_per_day": max_per_day, "min_per_day": min_per_day, "max_per_week": max_per_week,
            "time_limit": time_limit, "days": DAYS, "slots": SLOTS
        }

    def generate_timetable(self):
        """Kicks off the timetable generation process in a separate thread."""
        params = self._get_and_validate_params()
        if params is None:
            return

        self.status_var.set("Solving... The UI will remain responsive.")
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, "Solver is running in the background. Please wait...")
        self.root.update_idletasks()

        solver_thread = threading.Thread(target=self._execute_solver_thread, args=(params,), daemon=True)
        solver_thread.start()
        self.root.after(100, self._process_solver_queue)

    def _execute_solver_thread(self, params: Dict[str, Any]):
        """Runs the solver in a worker thread and puts the result in a queue."""
        try:
            schedule, room_assignments = solve_timetable(
                subjects=params["subjects"],
                num_batches=params["num_batches"],
                num_rooms=params["num_rooms"],
                max_per_day=params["max_per_day"],
                min_per_day=params["min_per_day"],
                max_per_week=params["max_per_week"],
                days=params["days"],
                slots=params["slots"],
                time_limit=params["time_limit"]
            )
            batch_names = [f"Batch {i+1}" for i in range(params["num_batches"])]
            result = {
                "status": "success",
                "data": {
                    "schedule": schedule,
                    "num_batches": params["num_batches"],
                    "num_rooms": params["num_rooms"],
                    "batch_names": batch_names,
                    "days": params["days"],
                    "slots": params["slots"],
                    "subjects": params["subjects"],
                    "room_assignments": room_assignments
                }
            }
            self.solver_queue.put(result)
        except Exception as e:
            self.solver_queue.put({"status": "error", "message": str(e)})

    def _process_solver_queue(self):
        """Checks the queue for solver results and updates the GUI."""
        try:
            result = self.solver_queue.get_nowait()
            if result["status"] == "error":
                messagebox.showerror("Solver Error / Infeasible", result["message"])
                self.status_var.set("Ready")
                self.text_area.delete(1.0, tk.END)
                return

            self.last_schedule_data = result["data"]
            self.display_results()
            self.status_var.set("Done. Timetable generated.")
        except queue.Empty:
            self.root.after(100, self._process_solver_queue)

    def _format_timetable_table(self) -> List[str]:
        """Formats the main schedule table as a list of strings."""
        out_lines = []
        schedule_data = self.last_schedule_data["schedule"]
        batch_names = self.last_schedule_data["batch_names"]
        days = self.last_schedule_data["days"]
        slots = self.last_schedule_data["slots"]
        num_days, num_slots_per_day = len(days), len(slots)
        total_slots = num_days * num_slots_per_day
        subjects_list = self.last_schedule_data["subjects"]

        max_subj_len = max(len(s[0]) for s in subjects_list) if subjects_list else 4
        slot_col_width = max(max_subj_len, len("Free")) + 2
        batch_col_name = "Batch/Slot"
        batch_col_width = max(len(bn) for bn in batch_names + [batch_col_name]) + 2
        total_col_header = "Total"
        total_col_width = len(total_col_header) + 2

        header_row = f"{batch_col_name:<{batch_col_width}}"
        for d in range(num_days):
            for k in range(num_slots_per_day):
                header_row += f"{f'{days[d]} {slots[k]}':<{slot_col_width}}"
            header_row += f"{total_col_header:<{total_col_width}}"
            if d < num_days - 1:
                header_row += " | "
        out_lines.append(header_row)
        out_lines.append("-" * len(header_row))

        for b, batch_name in enumerate(batch_names):
            row_str = f"{batch_names[b]:<{batch_col_width}}"
            weekly_total = 0
            for d in range(num_days):
                daily_total = 0
                for k in range(num_slots_per_day):
                    t = d * num_slots_per_day + k
                    cell = schedule_data.get(b, ["Free"] * total_slots)[t]
                    row_str += f"{cell:<{slot_col_width}}"
                    if cell != "Free":
                        daily_total += 1
                        weekly_total += 1
                row_str += f"{str(daily_total):<{total_col_width}}"
                if d < num_days - 1:
                    row_str += " | "
            out_lines.append(row_str)
            out_lines.append(f"{'':<{batch_col_width}}Weekly Total: {weekly_total}")
            out_lines.append("")
        return out_lines

    def _format_room_assignments(self) -> List[str]:
        """Formats the room assignment list as a list of strings."""
        out_lines = ["\n" + "="*40, "Room Assignments per Slot:", "="*40]
        days, slots = self.last_schedule_data["days"], self.last_schedule_data["slots"]
        batch_names = self.last_schedule_data["batch_names"]
        total_slots = len(days) * len(slots)

        for t in range(total_slots):
            d = t // len(slots)
            k = t % len(slots)
            slot_label = f"{days[d]} {slots[k]}"
            assigns = self.last_schedule_data["room_assignments"].get(t, [])
            if assigns:
                assigns_str = ", ".join((f"Room{r}-{batch_names[b]}" if r is not None else f"UNASSIGNED-{batch_names[b]}") for r, b in assigns)
                out_lines.append(f"{slot_label:<12}: {assigns_str}")
        return out_lines

    def _format_summary(self) -> List[str]:
        """Formats the summary statistics as a list of strings."""
        out_lines = ["\n" + "="*40, "Summary:", "="*40]
        schedule_data = self.last_schedule_data["schedule"]
        batch_names = self.last_schedule_data["batch_names"]
        num_rooms = self.last_schedule_data["num_rooms"]
        total_slots = len(self.last_schedule_data["days"]) * len(self.last_schedule_data["slots"])

        out_lines.append("Total scheduled hours per batch:")
        for b, batch_name in enumerate(batch_names):
            total = sum(1 for t in range(total_slots) if schedule_data.get(b, ["Free"] * total_slots)[t] != "Free")
            out_lines.append(f"- {batch_name}: {total} hours")

        full_count = 0
        for t in range(total_slots):
            filled = sum(1 for b in range(len(batch_names)) if schedule_data.get(b, ["Free"] * total_slots)[t] != "Free")
            if filled == num_rooms:
                full_count += 1
        out_lines.append(f"\nSlots with all {num_rooms} rooms occupied: {full_count} / {total_slots}")
        return out_lines

    def display_results(self):
        """Formats and displays all results in the text area."""
        table_lines = self._format_timetable_table()
        room_lines = self._format_room_assignments()
        summary_lines = self._format_summary()

        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, "\n".join(table_lines + room_lines + summary_lines))

def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        style.theme_use("clam")
    except Exception:
        pass
    app = TimetableApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
