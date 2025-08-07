# 🎓 Student Enrollment System (MS Access + VBA)

A lightweight and offline-capable **Student Enrollment Management System** developed using **Microsoft Access ** with **VBA (Visual Basic for Applications)**. Designed to help academic institutions streamline student data entry, ID generation, and reporting.

---

## 📌 Project Overview

This MS Access application allows institutions to:

- ✅ Store and manage student details: name, gender, DOB, contact info, course, etc.
- ✍️ Enter and update student records through a user-friendly form
- 🆔 Auto-generate **unique Student IDs**
- 🧾 Generate **printable reports** of student data with one click
- ⚙️ Enforce **input validation** using VBA logic
- 💾 Use an **offline** database with no external dependencies

---

## 🗃️ Database Structure

### Table: `Students`

| Field Name     | Data Type   | Description                         |
|----------------|-------------|-------------------------------------|
| `StudentID`    | AutoNumber  | Primary Key, auto-generated         |
| `FullName`     | Short Text  | Student's full name                 |
| `Gender`       | Short Text  | Gender (Male/Female)                |
| `DOB`          | Date/Time   | Date of Birth                       |
| `PhoneNumber`  | Short Text  | Contact number                      |
| `Email`        | Short Text  | Email address                       |
| `Course`       | Short Text  | Enrolled course                     |
| `EnrolmentDate`| Date/Time   | Date of enrollment                  |

---

## 🧾 Forms

### 🔹 `Std_Entry_Form`

- Add, edit, and view student data
- User interface includes:
  - Text boxes for input fields
  - Combo boxes for gender and course selection
- **VBA Features**:
  - Input validation
  - Record submission
  - Triggering report generation

---

## 📊 Reports

### 🔹 `rpt_StudentDetails`

- Auto-generated printable report of all students
- Triggered using a **command button** on the form
- Developed with **Access Reporting Tools** and **VBA macros**

---

## 📂 Project Contents

