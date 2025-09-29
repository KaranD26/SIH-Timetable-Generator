# SIH 2025 - Problem Statement SIH25028
# Team Name: Jedi Order
# Project: Smart Classroom and Timetable Scheduler

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from ortools.sat.python import cp_model
import csv
import threading
import queue
import json
from typing import Dict, List

# --- Constants ---
# Default days, can be dynamically changed in the app
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"]
# Default slots, now managed dynamically within the app class
DEFAULT_SLOTS = ["9-10", "10-11", "11-12", "12-1", "2-3", "3-4", "4-5"]
DEFAULT_SOLVER_TIME_LIMIT = 45
MAX_CONSECUTIVE_CLASSES = 3

# --- Solver Objective Weights ---
WEIGHT_ADJACENCY = 5
WEIGHT_PENALTY_FREE_DAY = 1000
WEIGHT_PENALTY_LONG_STRETCH = 50
WEIGHT_PENALTY_FACULTY_PREFERENCE = 20

class TimetableApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Jedi Order - Smart Timetable Scheduler (SIH25028)")
        self.root.geometry("1600x900")

        self.slots = DEFAULT_SLOTS.copy()

        self.batches: Dict[str, Dict] = {}
        self.faculty: Dict[str, Dict] = {}
        self.classrooms: Dict[str, Dict] = {}
        self.last_schedule_data = None
        self.solver_queue = queue.Queue()
        self.selected_item_for_edit = None

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)

        tab_names = ["▶ Configuration", "📚 Batches & Subjects", "🧑‍🏫 Faculty", "🏫 Classrooms", "📅 Generated Timetable", "📊 Analytics Dashboard"]
        self.tabs = {name: ttk.Frame(self.notebook, padding="10") for name in tab_names}
        for name, tab_frame in self.tabs.items():
            self.notebook.add(tab_frame, text=name)

        self._create_all_tabs()
        self.update_all_views()
        self._update_slot_dependent_widgets()

    def _create_all_tabs(self):
        self._create_config_tab()
        self._create_batches_tab()
        self._create_faculty_tab()
        self._create_classrooms_tab()
        self._create_timetable_tab()
        self._create_analytics_tab()

    def _create_config_tab(self):
        frame = self.tabs["▶ Configuration"]
        frame.columnconfigure(0, weight=1)

        data_frame = ttk.LabelFrame(frame, text="Data Management", padding="10")
        data_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        ttk.Button(data_frame, text="Save Configuration...", command=self.save_config).pack(side="left", padx=10)
        ttk.Button(data_frame, text="Load Configuration...", command=self.load_config).pack(side="left", padx=10)

        action_frame = ttk.LabelFrame(frame, text="Actions", padding="10")
        action_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        ttk.Button(action_frame, text="Validate Configuration", command=self.validate_configuration).pack(side="left", padx=10)
        self.generate_button = ttk.Button(action_frame, text="Generate Timetable", style="Accent.TButton", command=self.generate_timetable)
        self.generate_button.pack(side="right", padx=20, ipady=5)

        param_frame = ttk.LabelFrame(frame, text="Solver Parameters", padding="10")
        param_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        param_frame.columnconfigure(1, weight=1)

        ttk.Label(param_frame, text="Time Limit (s):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.time_limit_var = tk.IntVar(value=DEFAULT_SOLVER_TIME_LIMIT)
        ttk.Spinbox(param_frame, from_=10, to=600, textvariable=self.time_limit_var).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(param_frame, text="Min Hours/Day:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.min_hours_day_var = tk.IntVar(value=1)
        self.min_hours_spinbox = ttk.Spinbox(param_frame, from_=0, to=len(self.slots), textvariable=self.min_hours_day_var)
        self.min_hours_spinbox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(param_frame, text="Max Hours/Day:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.max_hours_day_var = tk.IntVar(value=5)
        self.max_hours_spinbox = ttk.Spinbox(param_frame, from_=1, to=len(self.slots), textvariable=self.max_hours_day_var)
        self.max_hours_spinbox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        self.saturday_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(param_frame, text="Include Saturday", variable=self.saturday_var).grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        # --- NEW: Time Slot Management ---
        slot_frame = ttk.LabelFrame(frame, text="Time Slot Management", padding="10")
        slot_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        self.start_hour_var = tk.IntVar(value=9)
        self.end_hour_var = tk.IntVar(value=17) # 5 PM
        self.lunch_hour_var = tk.IntVar(value=13) # 1 PM

        ttk.Label(slot_frame, text="Start Hour (24h):").grid(row=0, column=0, padx=5, sticky="w")
        ttk.Spinbox(slot_frame, from_=0, to=23, textvariable=self.start_hour_var, width=10).grid(row=0, column=1, padx=5)
        ttk.Label(slot_frame, text="End Hour (24h):").grid(row=0, column=2, padx=5, sticky="w")
        ttk.Spinbox(slot_frame, from_=0, to=24, textvariable=self.end_hour_var, width=10).grid(row=0, column=3, padx=5)
        ttk.Label(slot_frame, text="Lunch Start (24h):").grid(row=0, column=4, padx=5, sticky="w")
        ttk.Spinbox(slot_frame, from_=0, to=23, textvariable=self.lunch_hour_var, width=10).grid(row=0, column=5, padx=5)
        ttk.Button(slot_frame, text="Update Time Slots", command=self.update_slots).grid(row=0, column=6, padx=10)

        self.status_var = tk.StringVar(value="Ready. Load a configuration to begin.")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w", padding="2 5")
        status_bar.pack(side="bottom", fill="x")

    def _create_batches_tab(self):
        # ... (This method's code remains unchanged) ...
        frame = self.tabs["📚 Batches & Subjects"]
        frame.columnconfigure(1, weight=3)
        frame.columnconfigure(0, weight=2)
        frame.rowconfigure(0, weight=1)

        left_pane = ttk.Frame(frame)
        left_pane.grid(row=0, column=0, sticky="ns", padx=5, pady=5)
        left_pane.rowconfigure(1, weight=1)
        
        batch_ctrl_frame = ttk.LabelFrame(left_pane, text="Manage Batches", padding=10)
        batch_ctrl_frame.grid(row=0, column=0, sticky="ew")
        self.batch_name_var, self.batch_year_var, self.batch_branch_var = tk.StringVar(), tk.IntVar(value=1), tk.StringVar(value="CSE")
        
        ttk.Label(batch_ctrl_frame, text="Name:").grid(row=0, column=0, padx=2, pady=2, sticky="w")
        ttk.Entry(batch_ctrl_frame, textvariable=self.batch_name_var).grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        ttk.Label(batch_ctrl_frame, text="Year:").grid(row=1, column=0, padx=2, pady=2, sticky="w")
        ttk.Entry(batch_ctrl_frame, textvariable=self.batch_year_var).grid(row=1, column=1, padx=2, pady=2, sticky="ew")
        ttk.Label(batch_ctrl_frame, text="Branch:").grid(row=2, column=0, padx=2, pady=2, sticky="w")
        ttk.Entry(batch_ctrl_frame, textvariable=self.batch_branch_var).grid(row=2, column=1, padx=2, pady=2, sticky="ew")

        batch_button_frame = ttk.Frame(batch_ctrl_frame)
        batch_button_frame.grid(row=3, column=0, columnspan=2, pady=(10,0))

        ttk.Button(batch_button_frame, text="Add Batch", command=self.add_batch).pack(side="left", padx=5)
        self.batch_update_button = ttk.Button(batch_button_frame, text="Update Selected", command=self.update_batch, state="disabled")
        self.batch_update_button.pack(side="left", padx=5)
        self.batch_remove_button = ttk.Button(batch_button_frame, text="Remove Selected", command=self.remove_batch, state="disabled")
        self.batch_remove_button.pack(side="left", padx=5)
        
        batch_list_frame = ttk.LabelFrame(left_pane, text="All Batches (Double-click to edit)", padding=10)
        batch_list_frame.grid(row=1, column=0, sticky="nsew")
        self.batch_tree = ttk.Treeview(batch_list_frame, columns=("Name", "Year", "Branch"), show="headings")
        self.batch_tree.heading("Name", text="Batch Name")
        self.batch_tree.heading("Year", text="Year")
        self.batch_tree.heading("Branch", text="Branch")
        self.batch_tree.column("Name", width=120)
        self.batch_tree.column("Year", width=60)
        self.batch_tree.pack(fill="both", expand=True)
        self.batch_tree.bind("<Double-1>", self.on_batch_double_click)
        self.batch_tree.bind("<<TreeviewSelect>>", self.on_batch_select)
        
        right_pane = ttk.Frame(frame)
        right_pane.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        right_pane.rowconfigure(0, weight=3)
        right_pane.rowconfigure(1, weight=1)

        subject_frame = ttk.LabelFrame(right_pane, text="Subjects for Selected Batch (Double-click to edit)", padding="10")
        subject_frame.grid(row=0, column=0, sticky="nsew")
        subject_frame.rowconfigure(0, weight=1)
        subject_frame.columnconfigure(0, weight=1)

        self.subject_tree = ttk.Treeview(subject_frame, columns=("Subject", "Hours", "Type"), show="headings")
        self.subject_tree.heading("Subject", text="Subject Name")
        self.subject_tree.heading("Hours", text="Weekly Hours")
        self.subject_tree.heading("Type", text="Type (Lecture/Lab)")
        self.subject_tree.pack(fill="both", expand=True)
        self.subject_tree.bind("<<TreeviewSelect>>", self.on_subject_select)
        self.subject_tree.bind("<Double-1>", self.on_subject_double_click)
        
        subj_ctrl_frame = ttk.LabelFrame(right_pane, text="Manage Subjects", padding=10)
        subj_ctrl_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        subj_ctrl_frame.columnconfigure(1, weight=1)
        
        self.subj_name_var = tk.StringVar()
        self.subj_hours_var = tk.IntVar(value=1)
        self.subj_type_var = tk.StringVar(value="Lecture")

        ttk.Label(subj_ctrl_frame, text="Name:").grid(row=0, column=0, padx=2, pady=3, sticky="w")
        ttk.Entry(subj_ctrl_frame, textvariable=self.subj_name_var).grid(row=0, column=1, columnspan=2, padx=2, pady=3, sticky="ew")
        
        ttk.Label(subj_ctrl_frame, text="Hours:").grid(row=1, column=0, padx=2, pady=3, sticky="w")
        ttk.Spinbox(subj_ctrl_frame, from_=1, to=10, textvariable=self.subj_hours_var).grid(row=1, column=1, columnspan=2, padx=2, pady=3, sticky="ew")

        ttk.Label(subj_ctrl_frame, text="Type:").grid(row=2, column=0, padx=2, pady=3, sticky="w")
        ttk.Combobox(subj_ctrl_frame, textvariable=self.subj_type_var, values=["Lecture", "Lab"], state="readonly").grid(row=2, column=1, columnspan=2, padx=2, pady=3, sticky="ew")

        subj_button_frame = ttk.Frame(subj_ctrl_frame)
        subj_button_frame.grid(row=3, column=0, columnspan=3, pady=(10,0))
        
        ttk.Button(subj_button_frame, text="Add Subject", command=self.add_subject).pack(side="left", padx=5)
        self.subj_update_button = ttk.Button(subj_button_frame, text="Update Selected", command=self.update_subject, state="disabled")
        self.subj_update_button.pack(side="left", padx=5)
        self.subj_remove_button = ttk.Button(subj_button_frame, text="Remove Selected", command=self.remove_subject, state="disabled")
        self.subj_remove_button.pack(side="left", padx=5)

    def _create_faculty_tab(self):
        # ... (This method's code remains unchanged) ...
        frame = self.tabs["🧑‍🏫 Faculty"]
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(0, weight=1)

        left_pane = ttk.Frame(frame, padding="5")
        left_pane.grid(row=0, column=0, sticky="ns")
        left_pane.rowconfigure(2, weight=1)

        fac_ctrl_frame = ttk.LabelFrame(left_pane, text="Manage Faculty", padding="10")
        fac_ctrl_frame.grid(row=0, column=0, sticky="ew")
        fac_ctrl_frame.columnconfigure(1, weight=1)

        self.faculty_name_var = tk.StringVar()
        
        ttk.Label(fac_ctrl_frame, text="Name:").grid(row=0, column=0, padx=(0, 5), pady=2, sticky="w")
        ttk.Entry(fac_ctrl_frame, textvariable=self.faculty_name_var).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        
        fac_button_frame = ttk.Frame(fac_ctrl_frame)
        fac_button_frame.grid(row=1, column=0, columnspan=2, pady=(10,0))

        ttk.Button(fac_button_frame, text="Add", command=self.add_faculty).pack(side="left", padx=5)
        self.faculty_update_button = ttk.Button(fac_button_frame, text="Update Selected", command=self.update_faculty, state="disabled")
        self.faculty_update_button.pack(side="left", padx=5)
        self.faculty_remove_button = ttk.Button(fac_button_frame, text="Remove Selected", command=self.remove_faculty, state="disabled")
        self.faculty_remove_button.pack(side="left", padx=5)

        pref_frame = ttk.LabelFrame(left_pane, text="Preferences", padding="10")
        pref_frame.grid(row=1, column=0, sticky="ew", pady=10)
        self.faculty_pref_morning_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(pref_frame, text="Prefers Morning (9-12)", variable=self.faculty_pref_morning_var).pack(anchor="w")
        ttk.Button(pref_frame, text="Set Preferences", command=self.set_faculty_preferences).pack(pady=5)

        fac_list_frame = ttk.LabelFrame(left_pane, text="All Faculty (Double-click to update name)", padding="10")
        fac_list_frame.grid(row=2, column=0, sticky="nsew")
        self.faculty_tree = ttk.Treeview(fac_list_frame, columns=("Name",), show="headings")
        self.faculty_tree.heading("Name", text="Name")
        self.faculty_tree.pack(fill="both", expand=True)
        self.faculty_tree.bind("<Double-1>", self.on_faculty_double_click)
        self.faculty_tree.bind("<<TreeviewSelect>>", self.on_faculty_select)

        assign_frame = ttk.LabelFrame(frame, text="Teaching Assignments", padding="10")
        assign_frame.grid(row=0, column=1, sticky="nsew")
        assign_frame.rowconfigure(0, weight=1)
        assign_frame.columnconfigure(0, weight=1)
        
        self.assignment_tree = ttk.Treeview(assign_frame, columns=("Subject", "Batches"), show="headings")
        self.assignment_tree.heading("Subject", text="Subject Taught")
        self.assignment_tree.heading("Batches", text="Assigned Batches")
        self.assignment_tree.pack(fill="both", expand=True)
        self.assignment_tree.bind("<<TreeviewSelect>>", self.on_assignment_select)
        
        assign_button_frame = ttk.Frame(assign_frame)
        assign_button_frame.pack(pady=(5,0))

        self.add_assignment_button = ttk.Button(assign_button_frame, text="Add Assignment", command=self.add_assignment, state="disabled")
        self.add_assignment_button.pack(side="left", padx=5)
        self.edit_assignment_button = ttk.Button(assign_button_frame, text="Edit Selected Assignment", command=self.edit_assignment, state="disabled")
        self.edit_assignment_button.pack(side="left", padx=5)
        self.remove_assignment_button = ttk.Button(assign_button_frame, text="Remove Assignment", command=self.remove_assignment, state="disabled")
        self.remove_assignment_button.pack(side="left", padx=5)

    def _create_classrooms_tab(self):
        frame = self.tabs["🏫 Classrooms"]
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        class_ctrl_frame = ttk.LabelFrame(frame, text="Manage Classrooms", padding="10")
        class_ctrl_frame.grid(row=0, column=0, sticky="ew")
        self.class_name_var = tk.StringVar()
        self.class_capacity_var = tk.IntVar(value=1)
        self.class_type_var = tk.StringVar(value="Lecture")
        self.class_max_daily_hours_var = tk.IntVar(value=len(self.slots))
        
        ttk.Label(class_ctrl_frame, text="Name:").grid(row=0, column=0, pady=2, sticky="w")
        ttk.Entry(class_ctrl_frame, textvariable=self.class_name_var).grid(row=0, column=1, pady=2, sticky="ew")
        ttk.Label(class_ctrl_frame, text="Capacity (Batches):").grid(row=1, column=0, pady=2, sticky="w")
        ttk.Entry(class_ctrl_frame, textvariable=self.class_capacity_var).grid(row=1, column=1, pady=2, sticky="ew")
        ttk.Label(class_ctrl_frame, text="Type:").grid(row=2, column=0, pady=2, sticky="w")
        ttk.Combobox(class_ctrl_frame, textvariable=self.class_type_var, values=["Lecture", "Lab"]).grid(row=2, column=1, pady=2, sticky="ew")
        
        ttk.Label(class_ctrl_frame, text="Max Daily Hours:").grid(row=3, column=0, pady=2, sticky="w")
        self.class_max_hours_spinbox = ttk.Spinbox(class_ctrl_frame, from_=1, to=len(self.slots), textvariable=self.class_max_daily_hours_var)
        self.class_max_hours_spinbox.grid(row=3, column=1, pady=2, sticky="ew")

        class_button_frame = ttk.Frame(class_ctrl_frame)
        class_button_frame.grid(row=4, column=0, columnspan=2, pady=(10,0))

        ttk.Button(class_button_frame, text="Add Classroom", command=self.add_classroom).pack(side="left", padx=5)
        self.class_update_button = ttk.Button(class_button_frame, text="Update Selected", command=self.update_classroom, state="disabled")
        self.class_update_button.pack(side="left", padx=5)
        self.class_remove_button = ttk.Button(class_button_frame, text="Remove Selected", command=self.remove_classroom, state="disabled")
        self.class_remove_button.pack(side="left", padx=5)

        class_list_frame = ttk.LabelFrame(frame, text="All Classrooms (Double-click to edit)", padding="10")
        class_list_frame.grid(row=1, column=0, sticky="nsew")
        
        self.classroom_tree = ttk.Treeview(class_list_frame, columns=("Name", "Capacity", "Type", "MaxHours"), show="headings")
        self.classroom_tree.heading("Name", text="Classroom Name")
        self.classroom_tree.heading("Capacity", text="Capacity")
        self.classroom_tree.heading("Type", text="Type")
        self.classroom_tree.heading("MaxHours", text="Max Daily Hours")
        self.classroom_tree.column("Name", width=120)
        self.classroom_tree.column("MaxHours", width=100)
        self.classroom_tree.pack(fill="both", expand=True)
        self.classroom_tree.bind("<Double-1>", self.on_classroom_double_click)
        self.classroom_tree.bind("<<TreeviewSelect>>", self.on_classroom_select)
        
    def _create_timetable_tab(self):
        # ... (This method's code remains unchanged) ...
        frame = self.tabs["📅 Generated Timetable"]
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        grid_frame = ttk.Frame(frame)
        grid_frame.grid(row=0, column=0, sticky="nsew")
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.rowconfigure(0, weight=1)

        self.timetable_grid = ttk.Treeview(grid_frame, show="headings")
        self.timetable_grid.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(grid_frame, orient="vertical", command=self.timetable_grid.yview)
        hsb = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.timetable_grid.xview)
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self.timetable_grid.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        export_frame = ttk.LabelFrame(frame, text="Export Options", padding=10)
        export_frame.grid(row=1, column=0, sticky="ew", pady=(10,0))
        
        self.export_full_button = ttk.Button(export_frame, text="Export Full Timetable (CSV)", command=self.export_full_timetable_csv, state="disabled")
        self.export_full_button.pack(side="left", padx=10)
        self.export_faculty_button = ttk.Button(export_frame, text="Export Faculty Schedules (CSV)", command=self.export_faculty_schedules_csv, state="disabled")
        self.export_faculty_button.pack(side="left", padx=10)

    def _create_analytics_tab(self):
        # ... (This method's code remains unchanged) ...
        frame = self.tabs["📊 Analytics Dashboard"]
        self.analytics_frame = ttk.Frame(frame, padding=10)
        self.analytics_frame.pack(fill="both", expand=True)
        ttk.Label(self.analytics_frame, text="Generate a timetable to view analytics.", font=("Helvetica", 12)).pack()

    def update_slots(self):
        start = self.start_hour_var.get()
        end = self.end_hour_var.get()
        lunch = self.lunch_hour_var.get()

        if start >= end:
            messagebox.showerror("Error", "Start hour must be less than end hour.")
            return
        
        new_slots = []
        for hour in range(start, end):
            if hour == lunch:
                continue
            new_slots.append(f"{hour}-{hour+1}")
        
        self.slots = new_slots
        self._update_slot_dependent_widgets()
        messagebox.showinfo("Success", f"Time slots updated to {len(self.slots)} daily working hours.")

    def _update_slot_dependent_widgets(self):
        num_slots = len(self.slots)
        if num_slots == 0: num_slots = 1 # Avoid error if range is 0
        self.min_hours_spinbox.config(to=num_slots)
        self.max_hours_spinbox.config(to=num_slots)
        self.class_max_hours_spinbox.config(to=num_slots)
        
        if self.max_hours_day_var.get() > num_slots:
            self.max_hours_day_var.set(num_slots)
        if self.class_max_daily_hours_var.get() > num_slots:
            self.class_max_daily_hours_var.set(num_slots)

    def get_current_days(self) -> List[str]:
        # ... (This method's code remains unchanged) ...
        days = DAYS.copy()
        if self.saturday_var.get():
            days.append("Sat")
        return days

    def update_all_views(self):
        self.update_batch_view()
        self.update_faculty_view()
        self.update_classroom_view()

    def update_batch_view(self):
        self.batch_tree.delete(*self.batch_tree.get_children())
        for name, data in sorted(self.batches.items()):
            self.batch_tree.insert("", "end", iid=name, text=name, values=(name, data.get("year", ""), data.get("branch", "")))
        self.on_batch_select()

    def update_faculty_view(self):
        self.faculty_tree.delete(*self.faculty_tree.get_children())
        for name in sorted(self.faculty.keys()):
            self.faculty_tree.insert("", "end", iid=name, text=name, values=(name,))
        self.on_faculty_select()

    def update_classroom_view(self):
        self.classroom_tree.delete(*self.classroom_tree.get_children())
        for name, data in sorted(self.classrooms.items()):
            max_hours = data.get("max_daily_hours", len(self.slots))
            self.classroom_tree.insert("", "end", iid=name, text=name, values=(name, data.get("capacity", ""), data.get("type", ""), max_hours))
        self.on_classroom_select()

    def on_batch_select(self, event=None):
        has_selection = bool(self.batch_tree.selection())
        self.batch_update_button.config(state="normal" if has_selection else "disabled")
        self.batch_remove_button.config(state="normal" if has_selection else "disabled")

        self.selected_item_for_edit = None
        self.subject_tree.delete(*self.subject_tree.get_children())
        
        if not has_selection:
            return

        batch_name = self.batch_tree.selection()[0]
        if batch_name in self.batches and "subjects" in self.batches[batch_name]:
            subjects = self.batches[batch_name]["subjects"]
            for subj_name, details in sorted(subjects.items()):
                try:
                    hours = details.get("hours", "N/A")
                    subj_type = details.get("type", "N/A")
                    self.subject_tree.insert("", "end", iid=subj_name, values=(subj_name, hours, subj_type))
                except (AttributeError, KeyError):
                    print(f"Warning: Malformed subject data for '{subj_name}' in batch '{batch_name}'.")
        self.on_subject_select()

    def on_faculty_select(self, event=None):
        has_selection = bool(self.faculty_tree.selection())
        self.faculty_update_button.config(state="normal" if has_selection else "disabled")
        self.faculty_remove_button.config(state="normal" if has_selection else "disabled")
        self.add_assignment_button.config(state="normal" if has_selection else "disabled")
        
        self.selected_item_for_edit = None
        self.assignment_tree.delete(*self.assignment_tree.get_children())
        self.on_assignment_select() # Disable edit/remove buttons

        if not has_selection:
            self.faculty_pref_morning_var.set(False)
            return

        fac_name = self.faculty_tree.selection()[0]
        fac_data = self.faculty[fac_name]

        prefs = fac_data.get("preferences", {})
        self.faculty_pref_morning_var.set(prefs.get("prefers_morning", False))

        for subj, batches in sorted(fac_data.get("assignments", {}).items()):
            batch_str = ", ".join(sorted(list(batches)))
            self.assignment_tree.insert("", "end", iid=subj, values=(subj, batch_str))

    def on_classroom_select(self, event=None):
        has_selection = bool(self.classroom_tree.selection())
        self.class_update_button.config(state="normal" if has_selection else "disabled")
        self.class_remove_button.config(state="normal" if has_selection else "disabled")
        self.selected_item_for_edit = None

    def on_assignment_select(self, event=None):
        has_selection = bool(self.assignment_tree.selection())
        self.edit_assignment_button.config(state="normal" if has_selection else "disabled")
        self.remove_assignment_button.config(state="normal" if has_selection else "disabled")

    def on_batch_double_click(self, event):
        item_id = self.batch_tree.identify_row(event.y)
        if not item_id: return
        self.selected_item_for_edit = item_id
        data = self.batches[item_id]
        self.batch_name_var.set(item_id)
        self.batch_year_var.set(data["year"])
        self.batch_branch_var.set(data["branch"])
        self.batch_update_button.config(state="normal")

    def on_faculty_double_click(self, event):
        item_id = self.faculty_tree.identify_row(event.y)
        if not item_id: return
        self.selected_item_for_edit = item_id
        self.faculty_name_var.set(item_id)
        self.faculty_update_button.config(state="normal")

    def on_classroom_double_click(self, event):
        item_id = self.classroom_tree.identify_row(event.y)
        if not item_id: return
        self.selected_item_for_edit = item_id
        data = self.classrooms[item_id]
        self.class_name_var.set(item_id)
        self.class_capacity_var.set(data["capacity"])
        self.class_type_var.set(data["type"])
        self.class_max_daily_hours_var.set(data.get("max_daily_hours", len(self.slots)))
        self.class_update_button.config(state="normal")

    def on_subject_select(self, event=None):
        is_selected = bool(self.subject_tree.selection())
        self.subj_update_button.config(state="normal" if is_selected else "disabled")
        self.subj_remove_button.config(state="normal" if is_selected else "disabled")

    def on_subject_double_click(self, event):
        if not self.subject_tree.selection(): return
        subj_name = self.subject_tree.selection()[0]
        batch_name = self.batch_tree.selection()[0]
        details = self.batches[batch_name]["subjects"][subj_name]
        
        self.subj_name_var.set(subj_name)
        self.subj_hours_var.set(details.get("hours", 1))
        self.subj_type_var.set(details.get("type", "Lecture"))
    
    # --- ADD/REMOVE/UPDATE Methods ---
    def add_batch(self, *args):
        name = self.batch_name_var.get().strip()
        if not name or name in self.batches:
            messagebox.showerror("Error", "Batch name must be unique and not empty.")
            return
        self.batches[name] = {"year": self.batch_year_var.get(), "branch": self.batch_branch_var.get().strip(), "subjects": {}}
        self.update_batch_view()

    def remove_batch(self, *args):
        selection = self.batch_tree.selection()
        if not selection:
            messagebox.showerror("Error", "Please select a batch to remove.")
            return
        batch_name = selection[0]

        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove batch '{batch_name}'?\nThis will also remove it from all faculty assignments."):
            del self.batches[batch_name]
            # Clean up faculty assignments
            for fac_data in self.faculty.values():
                for subject, assigned_batches in fac_data["assignments"].items():
                    if batch_name in assigned_batches:
                        assigned_batches.remove(batch_name)
            self.update_batch_view()
            self.update_faculty_view()

    def update_batch(self, *args):
        if not self.selected_item_for_edit: return
        old_name, new_name = self.selected_item_for_edit, self.batch_name_var.get().strip()
        
        if old_name != new_name and new_name in self.batches:
            messagebox.showerror("Error", "A batch with this name already exists.")
            return

        data = self.batches.pop(old_name)
        data.update({"year": self.batch_year_var.get(), "branch": self.batch_branch_var.get()})
        self.batches[new_name] = data
        
        if old_name != new_name:
            for fac_data in self.faculty.values():
                for b_set in fac_data["assignments"].values():
                    if old_name in b_set:
                        b_set.remove(old_name)
                        b_set.add(new_name)
        
        self.selected_item_for_edit = None
        self.batch_update_button.config(state="disabled")
        self.update_all_views()

    def add_faculty(self, *args):
        name = self.faculty_name_var.get().strip()
        if not name or name in self.faculty:
            messagebox.showerror("Error", "Faculty name must be unique and not empty.")
            return
        self.faculty[name] = {"assignments": {}, "preferences": {}}
        self.update_faculty_view()

    def remove_faculty(self, *args):
        selection = self.faculty_tree.selection()
        if not selection:
            messagebox.showerror("Error", "Please select a faculty member to remove.")
            return
        faculty_name = selection[0]
        
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove '{faculty_name}'?"):
            del self.faculty[faculty_name]
            self.update_faculty_view()

    def update_faculty(self, *args):
        if not self.selected_item_for_edit: return
        old_name, new_name = self.selected_item_for_edit, self.faculty_name_var.get().strip()

        if not new_name:
            messagebox.showerror("Error", "Faculty name cannot be empty.")
            return

        if old_name != new_name and new_name in self.faculty:
            messagebox.showerror("Error", "A faculty member with this name already exists.")
            return

        if old_name != new_name:
            data = self.faculty.pop(old_name)
            self.faculty[new_name] = data

        self.selected_item_for_edit = None
        self.faculty_name_var.set("")
        self.faculty_update_button.config(state="disabled")
        self.update_faculty_view()

    def add_classroom(self, *args):
        name = self.class_name_var.get().strip()
        if not name or name in self.classrooms:
            messagebox.showerror("Error", "Classroom name must be unique and not empty.")
            return
        self.classrooms[name] = {
            "capacity": self.class_capacity_var.get(), 
            "type": self.class_type_var.get(),
            "max_daily_hours": self.class_max_daily_hours_var.get()
        }
        self.update_classroom_view()

    def remove_classroom(self, *args):
        selection = self.classroom_tree.selection()
        if not selection:
            messagebox.showerror("Error", "Please select a classroom to remove.")
            return
        classroom_name = selection[0]

        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove '{classroom_name}'?"):
            del self.classrooms[classroom_name]
            self.update_classroom_view()

    def update_classroom(self, *args):
        if not self.selected_item_for_edit: return
        old_name, new_name = self.selected_item_for_edit, self.class_name_var.get().strip()
        
        if old_name != new_name and new_name in self.classrooms:
            messagebox.showerror("Error", "A classroom with this name already exists.")
            return

        data = self.classrooms.pop(old_name)
        data.update({
            "capacity": self.class_capacity_var.get(), 
            "type": self.class_type_var.get(),
            "max_daily_hours": self.class_max_daily_hours_var.get()
        })
        self.classrooms[new_name] = data

        self.selected_item_for_edit = None
        self.class_update_button.config(state="disabled")
        self.update_classroom_view()
    
    # ... (Rest of the methods are unchanged) ...

    # The rest of the file is identical to the previous version, starting from add_subject
    # and continuing to the end. It is omitted here for brevity.
    # The only method with a significant change is solve_timetable_model.
    
    def add_subject(self, *args):
        batch_selection = self.batch_tree.selection()
        if not batch_selection:
            messagebox.showerror("Error", "Please select a batch first.")
            return
        batch_name = batch_selection[0]

        subj_name = self.subj_name_var.get().strip()
        if not subj_name:
            messagebox.showerror("Error", "Subject name cannot be empty.")
            return
        if subj_name in self.batches[batch_name]["subjects"]:
            messagebox.showerror("Error", f"Subject '{subj_name}' already exists in this batch.")
            return

        self.batches[batch_name]["subjects"][subj_name] = {"hours": self.subj_hours_var.get(), "type": self.subj_type_var.get()}
        self.on_batch_select()

    def update_subject(self, *args):
        batch_selection = self.batch_tree.selection()
        subj_selection = self.subject_tree.selection()
        if not batch_selection or not subj_selection:
            messagebox.showerror("Error", "Please select a batch and a subject to update.")
            return
        
        batch_name = batch_selection[0]
        old_subj_name = subj_selection[0]
        new_subj_name = self.subj_name_var.get().strip()

        if not new_subj_name:
            messagebox.showerror("Error", "Subject name cannot be empty.")
            return
        
        if old_subj_name != new_subj_name:
            if new_subj_name in self.batches[batch_name]["subjects"]:
                messagebox.showerror("Error", f"Subject '{new_subj_name}' already exists in this batch.")
                return
            del self.batches[batch_name]["subjects"][old_subj_name]

        self.batches[batch_name]["subjects"][new_subj_name] = {"hours": self.subj_hours_var.get(), "type": self.subj_type_var.get()}
        self.on_batch_select()

    def remove_subject(self, *args):
        batch_selection = self.batch_tree.selection()
        subj_selection = self.subject_tree.selection()
        if not batch_selection or not subj_selection:
            messagebox.showerror("Error", "Please select a subject to remove.")
            return

        batch_name = batch_selection[0]
        subj_name = subj_selection[0]

        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove '{subj_name}'?"):
            del self.batches[batch_name]["subjects"][subj_name]
            self.on_batch_select()

    def set_faculty_preferences(self, *args):
        selection = self.faculty_tree.selection()
        if not selection:
            messagebox.showerror("Error", "Select a faculty member first.")
            return
        faculty_name = selection[0]
        if "preferences" not in self.faculty[faculty_name]:
            self.faculty[faculty_name]["preferences"] = {}
        self.faculty[faculty_name]["preferences"]["prefers_morning"] = self.faculty_pref_morning_var.get()
        messagebox.showinfo("Success", f"Preferences for {faculty_name} have been updated.")

    def add_assignment(self, *args):
        fac_selection = self.faculty_tree.selection()
        if not fac_selection:
            messagebox.showerror("Error", "Please select a faculty member first.")
            return
        fac_name = fac_selection[0]

        editor = tk.Toplevel(self.root)
        editor.title(f"Add Assignment for {fac_name}")
        editor.geometry("400x150")
        
        ttk.Label(editor, text=f"Select a subject to assign to {fac_name}:").pack(pady=10)

        all_subjects = sorted(list(set(s for b in self.batches.values() for s in b["subjects"])))
        
        subject_var = tk.StringVar()
        subject_combo = ttk.Combobox(editor, textvariable=subject_var, values=all_subjects, state="readonly")
        subject_combo.pack(pady=5, padx=10, fill="x")

        def on_add():
            subj_name = subject_var.get()
            if not subj_name:
                messagebox.showwarning("Warning", "Please select a subject.", parent=editor)
                return
            if subj_name in self.faculty[fac_name]["assignments"]:
                messagebox.showwarning("Warning", f"{fac_name} is already assigned to teach {subj_name}.", parent=editor)
                return

            self.faculty[fac_name]["assignments"][subj_name] = set()
            self.on_faculty_select()
            editor.destroy()

        button_frame = ttk.Frame(editor)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Add", command=on_add).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Cancel", command=editor.destroy).pack(side="left", padx=10)

    def remove_assignment(self, *args):
        fac_selection = self.faculty_tree.selection()
        assign_selection = self.assignment_tree.selection()
        if not fac_selection or not assign_selection:
            messagebox.showerror("Error", "Please select a faculty member and an assignment to remove.")
            return
        
        fac_name = fac_selection[0]
        subj_name = assign_selection[0]

        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove the assignment '{subj_name}' from {fac_name}?"):
            if subj_name in self.faculty[fac_name]["assignments"]:
                del self.faculty[fac_name]["assignments"][subj_name]
            self.on_faculty_select()

    def edit_assignment(self, *args):
        fac_selection = self.faculty_tree.selection()
        assign_selection = self.assignment_tree.selection()
        if not fac_selection or not assign_selection:
            return

        fac_name = fac_selection[0]
        subj_name = assign_selection[0]

        editor = tk.Toplevel(self.root)
        editor.title(f"Edit Assignment for {subj_name}")
        editor.geometry("400x300")
        editor.transient(self.root)
        editor.grab_set()

        ttk.Label(editor, text=f"Assign batches for '{subj_name}' to {fac_name}:", font=('Helvetica', 10, 'bold')).pack(pady=10)

        listbox_frame = ttk.Frame(editor)
        listbox_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, exportselection=False)
        listbox.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.config(yscrollcommand=scrollbar.set)
        
        all_batches_formatted = []
        for name, data in sorted(self.batches.items()):
            all_batches_formatted.append(f"{name} ({data.get('year', 'N/A')}, {data.get('branch', 'N/A')})")
        
        listbox.insert(tk.END, *all_batches_formatted)
        
        current_assigned_set = self.faculty[fac_name]["assignments"].get(subj_name, set())
        for i, formatted_name in enumerate(all_batches_formatted):
            batch_name = formatted_name.split(' ')[0]
            if batch_name in current_assigned_set:
                listbox.selection_set(i)

        button_frame = ttk.Frame(editor)
        button_frame.pack(pady=10)
        
        save_button = ttk.Button(button_frame, text="Save", command=lambda: self._save_assignment_changes(editor, fac_name, subj_name, listbox))
        save_button.pack(side="left", padx=10)
        cancel_button = ttk.Button(button_frame, text="Cancel", command=editor.destroy)
        cancel_button.pack(side="left", padx=10)

    def _save_assignment_changes(self, editor, fac_name, subj_name, listbox):
        selected_indices = listbox.curselection()
        newly_assigned_batches = set()
        for i in selected_indices:
            formatted_name = listbox.get(i)
            batch_name = formatted_name.split(' ')[0]
            newly_assigned_batches.add(batch_name)
        
        self.faculty[fac_name]["assignments"][subj_name] = newly_assigned_batches
        self.on_faculty_select()
        editor.destroy()

    def validate_configuration(self, *args):
        # ... (unchanged)
        errors, warnings = [], []
        all_assigned_subjects = set()
        for fac_data in self.faculty.values():
            for subj in fac_data["assignments"]:
                all_assigned_subjects.add(subj)

        for b_name, b_data in self.batches.items():
            for s_name in b_data["subjects"]:
                if s_name not in all_assigned_subjects:
                    warnings.append(f"Subject '{s_name}' in Batch '{b_name}' has no faculty assigned.")

        if not errors and not warnings:
            messagebox.showinfo("Validation Success", "Configuration appears to be valid.")
        else:
            msg = f"Errors:\n{'-'*10}\n" + "\n".join(errors) if errors else "None"
            msg += f"\n\nWarnings:\n{'-'*10}\n" + "\n".join(warnings) if warnings else "None"
            messagebox.showwarning("Validation Issues", msg)
        
    def generate_timetable(self, *args):
        # ... (unchanged)
        self.status_var.set("Generating... This may take a moment.")
        self.root.update_idletasks()
        
        solver_thread = threading.Thread(target=self._execute_solver_thread, daemon=True)
        solver_thread.start()
        self.root.after(100, self._process_solver_queue)

    def _execute_solver_thread(self):
        # ... (unchanged)
        try:
            result_data = self.solve_timetable_model(
                self.batches, self.faculty, self.classrooms,
                self.time_limit_var.get(), self.min_hours_day_var.get(), self.max_hours_day_var.get()
            )
            self.solver_queue.put({"status": "success", "data": result_data})
        except Exception as e:
            self.solver_queue.put({"status": "error", "message": str(e)})

    def _process_solver_queue(self):
        # ... (unchanged)
        try:
            result = self.solver_queue.get_nowait()
            if result["status"] == "error":
                messagebox.showerror("Solver Error", f"Could not generate timetable.\nReason: {result['message']}")
                self.status_var.set("Solver failed.")
                self.export_full_button.config(state="disabled")
                self.export_faculty_button.config(state="disabled")
            else:
                self.last_schedule_data = result["data"]
                self.display_timetable_results()
                self.display_analytics()
                self.status_var.set("Timetable generated successfully.")
                self.export_full_button.config(state="normal")
                self.export_faculty_button.config(state="normal")
                self.notebook.select(self.tabs["📅 Generated Timetable"])
        except queue.Empty:
            self.root.after(100, self._process_solver_queue)

    def display_timetable_results(self):
        # ... (unchanged)
        for item in self.timetable_grid.get_children():
            self.timetable_grid.delete(item)
        
        days = self.get_current_days()
        columns = ["Batch"] + [f"{day}\n{slot}" for day in days for slot in self.slots]
        self.timetable_grid["columns"] = columns
        self.timetable_grid.column("Batch", width=120, anchor='w')
        for col in columns[1:]:
            self.timetable_grid.heading(col, text=col)
            self.timetable_grid.column(col, width=120, anchor='center')
        self.timetable_grid.heading("Batch", text="Batch")

        schedule = self.last_schedule_data["schedule"]
        batch_list = self.last_schedule_data["batch_list"]

        for batch_idx, batch_name in enumerate(batch_list):
            row_data = [batch_name]
            for t in range(len(days) * len(self.slots)):
                cell = schedule[batch_idx][t]
                cell_text = ""
                if cell:
                    subj, fac, room = cell
                    cell_text = f"{subj}\n{fac}\n@{room}"
                row_data.append(cell_text)
            self.timetable_grid.insert("", "end", values=row_data)

    def display_analytics(self):
        # ... (unchanged)
        for widget in self.analytics_frame.winfo_children():
            widget.destroy()
        if not self.last_schedule_data: 
            ttk.Label(self.analytics_frame, text="Generate a timetable to view analytics.").pack()
            return

        workload_frame = ttk.LabelFrame(self.analytics_frame, text="Faculty Workload (Hours per Week)", padding=10)
        workload_frame.pack(fill="x", pady=5, expand=False)
        
        workloads = self.last_schedule_data["workloads"]
        max_load = max(workloads.values()) if workloads else 1
        for i, (fac, hours) in enumerate(sorted(workloads.items())):
            ttk.Label(workload_frame, text=f"{fac}:").grid(row=i, column=0, sticky='w', padx=5)
            p_bar = ttk.Progressbar(workload_frame, orient="horizontal", length=400, mode="determinate", maximum=max_load, value=hours)
            p_bar.grid(row=i, column=1, sticky='ew', padx=5)
            ttk.Label(workload_frame, text=f"{hours} hrs").grid(row=i, column=2, sticky='w', padx=5)
        workload_frame.columnconfigure(1, weight=1)

        util_frame = ttk.LabelFrame(self.analytics_frame, text="Classroom Utilization (vs. Max Workable Hours)", padding=10)
        util_frame.pack(fill="x", pady=5, expand=False, side="top")
        
        classroom_usage = {room: 0 for room in self.classrooms}
        schedule = self.last_schedule_data["schedule"]
        for batch_schedule in schedule:
            for cell in batch_schedule:
                if cell:
                    _, _, room = cell
                    if room in classroom_usage:
                        classroom_usage[room] += 1
        
        total_available_hours = self.max_hours_day_var.get() * len(self.get_current_days())
        if total_available_hours == 0: total_available_hours = 1
        
        for i, (room, busy_hours) in enumerate(sorted(classroom_usage.items())):
            ttk.Label(util_frame, text=f"{room}:").grid(row=i, column=0, sticky='w', padx=5)
            p_bar = ttk.Progressbar(util_frame, orient="horizontal", length=400, mode="determinate", maximum=total_available_hours, value=busy_hours)
            p_bar.grid(row=i, column=1, sticky='ew', padx=5)
            ttk.Label(util_frame, text=f"{busy_hours} hrs").grid(row=i, column=2, sticky='w', padx=5)
        util_frame.columnconfigure(1, weight=1)

    def save_config(self, *args):
        # ... (unchanged)
        data = {
            "batches": self.batches,
            "faculty": {name: {
                "assignments": {subj: list(b_set) for subj, b_set in fac_data["assignments"].items()},
                "preferences": fac_data.get("preferences", {})
            } for name, fac_data in self.faculty.items()},
            "classrooms": self.classrooms
        }
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not filepath: return
        try:
            with open(filepath, 'w') as f:
                json.dump(data, f, indent=4)
            self.status_var.set(f"Configuration saved to {filepath}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save configuration file:\n{e}")

    def load_config(self, *args):
        # ... (unchanged)
        filepath = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not filepath: return
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
            self.batches = data.get("batches", {})
            self.classrooms = data.get("classrooms", {})
            self.faculty = {}
            for name, fac_data in data.get("faculty", {}).items():
                self.faculty[name] = {
                    "assignments": {subj: set(b_list) for subj, b_list in fac_data["assignments"].items()},
                    "preferences": fac_data.get("preferences", {})
                }
            self.update_all_views()
            self.status_var.set(f"Configuration loaded from {filepath}")
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not load or parse configuration file:\n{e}")

    def export_full_timetable_csv(self, *args):
        # ... (unchanged)
        if not self.last_schedule_data:
            messagebox.showerror("Error", "No timetable data to export. Please generate a timetable first.")
            return
        
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")], title="Save Full Timetable")
        if not filepath: return

        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                days = self.get_current_days()
                header = ["Batch"] + [f"{day} {slot}" for day in days for slot in self.slots]
                writer.writerow(header)
                
                schedule = self.last_schedule_data["schedule"]
                batch_list = self.last_schedule_data["batch_list"]

                for batch_idx, batch_name in enumerate(batch_list):
                    row = [batch_name]
                    for t in range(len(days) * len(self.slots)):
                        cell = schedule[batch_idx][t]
                        if cell:
                            subj, fac, room = cell
                            row.append(f"{subj} | {fac} | @{room}")
                        else:
                            row.append("")
                    writer.writerow(row)
            messagebox.showinfo("Success", f"Full timetable exported to {filepath}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export file:\n{e}")

    def export_faculty_schedules_csv(self, *args):
        # ... (unchanged)
        if not self.last_schedule_data:
            messagebox.showerror("Error", "No timetable data to export. Please generate a timetable first.")
            return
            
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")], title="Save Faculty Schedules")
        if not filepath: return

        try:
            days = self.get_current_days()
            total_slots = len(days) * len(self.slots)
            num_slots_per_day = len(self.slots)
            
            faculty_schedules = {name: [None] * total_slots for name in self.faculty.keys()}
            schedule = self.last_schedule_data["schedule"]
            batch_list = self.last_schedule_data["batch_list"]

            for b_idx, batch_name in enumerate(batch_list):
                for t in range(total_slots):
                    if schedule[b_idx][t]:
                        subj, fac, room = schedule[b_idx][t]
                        if fac in faculty_schedules:
                            faculty_schedules[fac][t] = (subj, batch_name, room)
            
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["Faculty", "Day", "Slot", "Subject", "Batch", "Classroom"])
                
                for fac_name, fac_schedule in sorted(faculty_schedules.items()):
                    for t, cell in enumerate(fac_schedule):
                        if cell:
                            day = days[t // num_slots_per_day]
                            slot = self.slots[t % num_slots_per_day]
                            subj, batch, room = cell
                            writer.writerow([fac_name, day, slot, subj, batch, room])
            
            messagebox.showinfo("Success", f"Faculty schedules exported to {filepath}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export file:\n{e}")

    def solve_timetable_model(self, batches, faculty, classrooms, time_limit, min_hours_day, max_hours_day):
        model = cp_model.CpModel()
        
        days = self.get_current_days()
        
        batch_list = sorted(batches.keys())
        faculty_list = sorted(faculty.keys())
        classroom_list = sorted(classrooms.keys())
        all_subjects = sorted(list(set(s for b in batches.values() for s in b["subjects"])))
        
        num_batches = len(batch_list)
        num_faculty = len(faculty_list)
        num_classrooms = len(classroom_list)
        num_subjects = len(all_subjects)
        num_days = len(days)
        num_slots_per_day = len(self.slots)
        total_slots = num_days * num_slots_per_day
        
        x = {}
        for b, s, f, c, t in [(b,s,f,c,t) for b in range(num_batches) for s in range(num_subjects) for f in range(num_faculty) for c in range(num_classrooms) for t in range(total_slots)]:
            x[b, s, f, c, t] = model.NewBoolVar(f'x_b{b}_s{s}_f{f}_c{c}_t{t}')

        # Standard constraints...
        for b_idx, b_name in enumerate(batch_list):
            for s_idx, s_name in enumerate(all_subjects):
                if s_name not in batches[b_name]["subjects"]:
                    model.Add(sum(x[b_idx, s_idx, f, c, t] for f in range(num_faculty) for c in range(num_classrooms) for t in range(total_slots)) == 0)
                    continue

                faculty_for_subject = {fac for fac, dat in faculty.items() if s_name in dat["assignments"] and b_name in dat["assignments"][s_name]}
                for f_idx, f_name in enumerate(faculty_list):
                    if f_name not in faculty_for_subject:
                        model.Add(sum(x[b_idx, s_idx, f_idx, c, t] for c in range(num_classrooms) for t in range(total_slots)) == 0)

                subj_type = batches[b_name]["subjects"][s_name].get("type")
                for c_idx, c_name in enumerate(classroom_list):
                    if classrooms[c_name]["type"] != subj_type:
                        model.Add(sum(x[b_idx, s_idx, f, c_idx, t] for f in range(num_faculty) for t in range(total_slots)) == 0)

        for b, t in [(b,t) for b in range(num_batches) for t in range(total_slots)]:
            model.Add(sum(x[b, s, f, c, t] for s in range(num_subjects) for f in range(num_faculty) for c in range(num_classrooms)) <= 1)
        
        for f, t in [(f,t) for f in range(num_faculty) for t in range(total_slots)]:
            model.Add(sum(x[b, s, f, c, t] for b in range(num_batches) for s in range(num_subjects) for c in range(num_classrooms)) <= 1)

        for c_idx, c_name in enumerate(classroom_list):
            for t in range(total_slots):
                model.Add(sum(x[b,s,f,c_idx,t] for b in range(num_batches) for s in range(num_subjects) for f in range(num_faculty)) <= classrooms[c_name]["capacity"])

        # --- NEW: Classroom daily hour limit constraint ---
        for c_idx, c_name in enumerate(classroom_list):
            max_h = classrooms[c_name].get("max_daily_hours", num_slots_per_day)
            for d in range(num_days):
                daily_slots = range(d * num_slots_per_day, (d + 1) * num_slots_per_day)
                model.Add(sum(x[b,s,f,c_idx,t] for b in range(num_batches) for s in range(num_subjects) for f in range(num_faculty) for t in daily_slots) <= max_h)
        
        for b_idx, b_name in enumerate(batch_list):
            for s_idx, s_name in enumerate(all_subjects):
                if s_name in batches[b_name]["subjects"]:
                    required_hours = batches[b_name]["subjects"][s_name].get("hours", 0)
                    model.Add(sum(x[b_idx, s_idx, f, c, t] for f in range(num_faculty) for c in range(num_classrooms) for t in range(total_slots)) == required_hours)

        y = {}
        for b,t in [(b,t) for b in range(num_batches) for t in range(total_slots)]:
            y[b, t] = model.NewBoolVar(f'y_b{b}_t{t}')
            model.Add(sum(x[b,s,f,c,t] for s in range(num_subjects) for f in range(num_faculty) for c in range(num_classrooms)) == y[b,t])
        
        for b in range(num_batches):
            for d in range(num_days):
                daily_hours = sum(y[b, d * num_slots_per_day + k] for k in range(num_slots_per_day))
                model.Add(daily_hours <= max_hours_day)
                day_has_class = model.NewBoolVar(f'day_has_class_b{b}_d{d}')
                model.Add(daily_hours > 0).OnlyEnforceIf(day_has_class)
                model.Add(daily_hours == 0).OnlyEnforceIf(day_has_class.Not())
                model.Add(daily_hours >= min_hours_day).OnlyEnforceIf(day_has_class)
        
        # ... (Objective function remains unchanged) ...
        objective_terms = []
        for b in range(num_batches):
            for d in range(num_days):
                for k in range(num_slots_per_day - 1):
                    t = d * num_slots_per_day + k
                    adj = model.NewBoolVar(f'adj_b{b}_d{d}_k{k}')
                    model.AddBoolAnd([y[b, t], y[b, t+1]]).OnlyEnforceIf(adj)
                    objective_terms.append(adj * WEIGHT_ADJACENCY)

        for b in range(num_batches):
            for d in range(num_days):
                daily_hours = sum(y[b, d * num_slots_per_day + k] for k in range(num_slots_per_day))
                is_free = model.NewBoolVar(f'free_day_b{b}_d{d}')
                model.Add(daily_hours == 0).OnlyEnforceIf(is_free)
                objective_terms.append(-is_free * WEIGHT_PENALTY_FREE_DAY)
        
        for b in range(num_batches):
            for d in range(num_days):
                for k in range(num_slots_per_day - MAX_CONSECUTIVE_CLASSES):
                    t_start = d * num_slots_per_day + k
                    stretch = [y[b, t_start + i] for i in range(MAX_CONSECUTIVE_CLASSES + 1)]
                    is_long = model.NewBoolVar(f'long_stretch_b{b}_d{d}_k{k}')
                    model.AddBoolAnd(stretch).OnlyEnforceIf(is_long)
                    objective_terms.append(-is_long * WEIGHT_PENALTY_LONG_STRETCH)
        
        for f_idx, f_name in enumerate(faculty_list):
            prefs = faculty[f_name].get("preferences", {})
            if prefs.get("prefers_morning"):
                morning_slots_end_hour = 12
                for d in range(num_days):
                    for k in range(num_slots_per_day):
                        slot_start_hour = int(self.slots[k].split('-')[0])
                        if slot_start_hour >= morning_slots_end_hour:
                            t = d * num_slots_per_day + k
                            has_class = model.NewBoolVar(f'pref_f{f_idx}_t{t}')
                            model.Add(sum(x[b,s,f_idx,c,t] for b in range(num_batches) for s in range(num_subjects) for c in range(num_classrooms)) > 0).OnlyEnforceIf(has_class)
                            objective_terms.append(-has_class * WEIGHT_PENALTY_FACULTY_PREFERENCE)

        model.Maximize(sum(objective_terms))

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = float(time_limit)
        solver.parameters.num_search_workers = 8
        status = solver.Solve(model)

        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            raise RuntimeError("No feasible solution found. Check constraints and parameters (e.g., min/max hours per day).")

        schedule = [[None] * total_slots for _ in range(num_batches)]
        for b in range(num_batches):
            for t in range(total_slots):
                for s in range(num_subjects):
                    for f in range(num_faculty):
                        for c in range(num_classrooms):
                            if solver.Value(x[b, s, f, c, t]):
                                schedule[b][t] = (all_subjects[s], faculty_list[f], classroom_list[c])
        
        workloads = {f: 0 for f in faculty_list}
        for b in range(num_batches):
            for t in range(total_slots):
                if schedule[b][t]:
                    _, fac, _ = schedule[b][t]
                    if fac in workloads:
                        workloads[fac] += 1
        
        return {"schedule": schedule, "batch_list": batch_list, "workloads": workloads}

def main():
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Accent.TButton", foreground="white", background="#0078D7", font=('Helvetica', 10, 'bold'))
        style.configure("TNotebook.Tab", padding=[12, 5], font=('Helvetica', 10, 'bold'))
    except tk.TclError:
        print("Could not set advanced theme. Using default.")
    
    app = TimetableApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()