# EXAM_SCHEDULER

*Streamlining Exams for Seamless Student Success*

![last commit](https://img.shields.io/github/last-commit/farisahmed123/Exam_Scheduler)
![python](https://img.shields.io/badge/python-100.0%25-blue)
![languages](https://img.shields.io/badge/languages-1-green)

Built with the tools and technologies:

![Python](https://img.shields.io/badge/Python-3776AB?logo=python&logoColor=white)

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Usage](#usage)
- [Project Structure](#project-structure)
- [Algorithm & Approach](#algorithm--approach)
- [Output Files](#output-files)
- [Configuration](#configuration)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgments](#acknowledgments)

---

## Overview

The **Exam Scheduler** is a Graph Theory-based application designed to automate exam scheduling for educational institutions. It creates clash-free, optimized exam timetables while respecting constraints such as student conflicts, slot capacity, and course combinations.

This project was developed for **FAST NUCES Peshawar Campus** to schedule Sessional 1 Exams across 3 days with 6 time slots per day, ensuring no student has overlapping exams and minimizing student stress by reducing consecutive exams and multiple exams on the same day.

---

## Features

✅ **Clash-Free Scheduling**: Ensures no student has two exams at the same time  
✅ **Combined Course Support**: Schedules specified courses in the same slot  
✅ **Capacity Management**: Respects maximum student limit per slot (500 students)  
✅ **Smart Optimization**: Minimizes students with 3+ papers on the same day  
✅ **Consecutive Exam Reduction**: Reduces back-to-back exams for students  
✅ **Flexible Input**: Works with Excel files with customizable column names  
✅ **Detailed Reports**: Generates comprehensive statistics and analysis  
✅ **Excel Output**: Professional timetable in Excel format with multiple views

---

## Getting Started

### Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

### Installation

1. Clone the repository:
```bash
git clone https://github.com/farisahmed123/Exam_Scheduler.git
cd Exam_Scheduler
```

2. Install required packages:
```bash
pip install pandas openpyxl
```

Or use the requirements file (if available):
```bash
pip install -r requirements.txt
```

---

## Usage

1. **Prepare your data file**: Place your Excel file with student enrollment data in the project directory. The file should contain:
   - Student ID / Roll Number
   - Course Code
   - Section
   - Course/Subject Name

2. **Update the filename** in `exam_scheduler.py` (line 414):
```python
scheduler = ExamScheduler("Student Data.xlsx")  # Change to your filename
```

3. **Run the scheduler**:
```bash
python exam_scheduler.py
```

4. **Check the output**: The program generates two files:
   - `exam_schedule.xlsx` - Complete exam timetable
   - `schedule_report.txt` - Detailed statistics and analysis

---

## Project Structure

```
Exam_Scheduler/
├── exam_scheduler.py      # Main scheduling algorithm
├── Student Data.xlsx      # Input: Student enrollment data
├── exam_schedule.xlsx     # Output: Generated exam schedule
├── schedule_report.txt    # Output: Statistics and analysis
├── .gitignore            # Git ignore file
└── README.md             # Project documentation
```

---

## Algorithm & Approach

The scheduler uses a **Graph Coloring Algorithm** based on Graph Theory principles:

### 1. **Conflict Graph Construction**
- Each course is represented as a node
- Edges connect courses with common students (conflicts)
- Courses with edges cannot be scheduled simultaneously

### 2. **Combined Course Detection**
- Reads Excel columns to identify courses that must be in the same slot
- Extracts course codes using regex patterns
- Groups courses that should be scheduled together

### 3. **Greedy Graph Coloring with Optimization**
- Courses prioritized by:
  - Number of conflicts (graph degree)
  - Number of enrolled students
- Each "color" represents a time slot (day + slot)
- Penalty-based selection minimizes:
  - Students with 3+ papers on same day
  - Consecutive exams without breaks

### 4. **Constraints Satisfied**
- ✅ No scheduling conflicts (clash-free)
- ✅ Maximum 500 students per slot
- ✅ All exams within 3 days × 6 slots = 18 total slots
- ✅ Combined courses scheduled together

---

## Output Files

### 1. `exam_schedule.xlsx`
Excel file with two sheets:
- **Time Table**: Slot-wise view showing all courses in each time slot
- **Course List**: Course-wise view showing when each course is scheduled

### 2. `schedule_report.txt`
Detailed report containing:
- **Overview**: Total students, courses, and slots
- **Slot Load Analysis**: Number of students in each slot
- **Student Impact Analysis**:
  - Students with 3, 4, 5, or 6 papers on one day
  - Students with 3, 4, 5, or 6 consecutive papers
- **Approach Description**: Methodology used

---

## Configuration

You can modify the following parameters in `exam_scheduler.py`:

```python
class ExamScheduler:
    def __init__(self, excel_file):
        self.DAYS = 3                # Number of exam days
        self.SLOTS_PER_DAY = 6       # Slots per day
        self.MAX_CAPACITY = 500      # Max students per slot
```

### Time Slots
Current configuration (1-hour slots):
- Slot 1: 08:00 - 09:00
- Slot 2: 09:00 - 10:00
- Slot 3: 10:00 - 11:00
- Slot 4: 11:00 - 12:00
- Slot 5: 12:00 - 13:00
- Slot 6: 13:00 - 14:00

---

## Contributing

Contributions are welcome! If you'd like to improve the scheduler:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/improvement`)
3. Commit your changes (`git commit -m 'Add improvement'`)
4. Push to the branch (`git push origin feature/improvement`)
5. Open a Pull Request

---

## License

This project is open source and available under the [MIT License](LICENSE).

---

## Acknowledgments

- Developed for **FAST NUCES Peshawar Campus**
- Graph Theory concepts applied for optimal scheduling
- Thanks to the pandas and openpyxl libraries for data handling

---

## Contact

For questions or support, please open an issue on GitHub.

**Repository**: [https://github.com/farisahmed123/Exam_Scheduler](https://github.com/farisahmed123/Exam_Scheduler)

---

*Made with ❤️ for better exam scheduling*
