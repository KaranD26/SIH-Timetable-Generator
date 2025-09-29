# JEDI ORDER - SIH'25
# Automated Timetable Generator (SIH Problem #25028)

This is a functional desktop prototype for the Smart India Hackathon, designed to generate optimized academic timetables. It is built with Python, Tkinter (for the GUI), and Google's OR-Tools for constraint solving.

The final goal is to develop a full-stack **web application** with a polished, interactive user interface.

## Core Features (Prototype)
* GUI for managing all core data: **Batches, Subjects, Faculty, and Classrooms**.
* **Dynamic Time Slot Configuration**, including setting start/end times and lunch breaks.
* Solves hard constraints like **faculty availability, classroom capacity, subject-classroom type matching (Lab/Lecture), and weekly hour requirements**.
* Optimizes schedules based on soft constraints to **minimize gaps, avoid long class stretches, and respect faculty preferences**.
* Exports the final comprehensive timetable and individual faculty schedules to **CSV files**.
* Includes a basic **Analytics Dashboard** to visualize faculty workload and classroom utilization.

## How to Run
1.  **Install prerequisites:** `pip install ortools`
2.  **Run the script:** `python TimetableGenerator.py` or double-click `Execute.bat`

## Roadmap & Future Features
The planned web application will include:
* **Dynamic Timetable Interaction:** Real-time conflict checking and the ability for admins to make manual drag-and-drop adjustments.
* **Role-Based Dashboards:** Separate, customized timetable views for administrators, faculty, and batches.
* **Database Integration:** Centralized management of all institutional data (subjects, rooms, etc.).
* **Advanced Constraints:** Support for electives, faculty availability, and room-specific requirements.
* **User Accounts & Authentication.**