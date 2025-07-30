# Dorm Attendance Form Generator

A Python script to generate weekly dormitory attendance/check-in forms in `.docx` format for up to 18 weeks, based on a Word template.

## 📋 Features

- Auto-fills weekly dates starting from a given start date
- Customizes title with underline formatting for the week number
- Modifies the first cell of the Word table to inject a new heading
- Uses Microsoft Word `.docx` templates and maintains original styling
- Outputs 18 weekly attendance forms with correct filenames

## 📁 Template Requirements

The script requires a `.docx` template containing a table:
- The first row and cell (`cell(0, 0)`) will be overwritten by the week title.
- The second to eighth rows in the first column (`cell(i+2, 0)`) will be filled with daily dates (Monday to Sunday).

## 🛠️ Dependencies

Install required packages via pip:

```bash
pip install python-docx
