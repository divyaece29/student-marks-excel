# Student Marks Excel Generator

A Python script that takes student marks as input and automatically generates a formatted Excel report.

## Features
- Takes student name and marks as user input
- Calculates total, average, and grade automatically
- Formatted header row with bold text and blue background
- Two sheets — All Students and Toppers (average above 80)

## Grade Criteria
- 90 and above → A+
- 80 to 89 → A
- 70 to 79 → B
- 60 to 69 → C
- Below 60 → F

## Libraries Used
- openpyxl

## How to Run
Install the library:
pip install openpyxl

Then run:
python student_marks.py

## Output
Generates student_marks.xlsx with two sheets
