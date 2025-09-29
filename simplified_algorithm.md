# The Algorithm, Simplified 🧠

Think of the scheduling problem as solving a giant, multi-dimensional **Sudoku puzzle**. The computer's goal is to fill every time slot according to a set of rules and preferences.

---

## 🧩 The Puzzle Pieces

The algorithm first identifies all the components it needs to schedule:
* **Batches** (Student Groups)
* **Subjects** (Math, Physics, etc.)
* **Faculty** (The teachers)
* **Classrooms** (Lecture halls, labs)
* **Time Slots** (e.g., "Monday 9-10 AM")

---

## 🤔 The Core Question

For every single time slot, the algorithm makes one core decision:
> **Who** teaches **What** to **Which Batch** in **Which Room**?

It represents this as a massive checklist of `YES/NO` options for every possible combination.

---

## 📜 The Unbreakable Rules (Hard Constraints)

Next, it applies strict rules that **cannot be broken**:
* **No Double-Booking:** A teacher, a classroom, or a batch can only be in one place at a time.
* **Meet The Quota:** Every subject must be taught for its required number of hours per week.
* **Right Tool for the Job:** Lab classes must happen in lab rooms, and lectures in lecture halls.
* **Teacher Assignments:** A subject must be taught by a professor who is actually assigned to it.

---

## ⭐ The "High Score" Goals (The Objective)

Finally, among all the *possible* timetables, it tries to find the *best* one by aiming for a high score.

* **It Likes 👍:**
    * **Back-to-back classes** to minimize gaps.
* **It Dislikes 👎:**
    * **Long, tiring stretches** of 3+ consecutive classes.
    * **Scheduling teachers** outside of their preferred times (e.g., morning).
    * **Giving a batch an empty day** in the middle of the week.

The final timetable is the one that follows all the rules and gets the highest possible score.