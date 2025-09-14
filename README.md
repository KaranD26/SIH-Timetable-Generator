# JEDI ORDER - SIH'25
# Automated Timetable Generator (SIH Problem #25028)

This is a functional desktop prototype for the Smart India Hackathon, designed to generate optimized academic timetables. It is built with Python, Tkinter (for the GUI), and Google's OR-Tools for constraint solving.

The final goal is to develop a full-stack **web application** with a polished, interactive user interface.

## Core Features (Prototype)
* GUI for setting parameters (batches, rooms, subjects).
* Solves hard constraints like room/faculty limits and weekly hours.
* Optimizes the schedule to reduce gaps and maximize resource usage.
* Exports the final timetable to a CSV file.

## How to Run
1.  **Install prerequisites:** `pip install ortools`
2.  **Run the script:** `python TT_prototype.py`

## Roadmap & Future Features
The planned web application will include:
* **Dynamic Timetable Interaction:** Real-time conflict checking and the ability for admins to make manual drag-and-drop adjustments.
* **Role-Based Dashboards:** Separate, customized timetable views for administrators, faculty, and batches.
* **Database Integration:** Centralized management of all institutional data (subjects, rooms, etc.).
* **Advanced Constraints:** Support for electives, faculty availability, and room-specific requirements.
* **User Accounts & Authentication.**