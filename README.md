## Dormitory Safety and Fire Confirmation Form Generator
📅 For Weeks 1–18 | 🔒 Safety Confirmation Oriented | 🇨🇳 SDU Weihai
A Python script to generate weekly dormitory attendance/check-in forms in `.docx` format for up to 18 weeks, based on a Word template.

## NOTICE
This Python script automates the generation of dormitory safety and fire prevention confirmation forms for Weeks 1–18 at Shandong University, Weihai.

Each form corresponds to Week n and automatically includes confirmation entries for the period spanning Week n−4 to Week n.

As these forms are designed for safety confirmation rather than incident reporting, there is no need to retroactively submit unresolved issues. This ensures temporal consistency and aligns with the university's safety documentation protocol.

The project is open-source and available on GitHub:
🔗 https://github.com/Karcen/SDUWH-Dorm-Attendance-Form-Generator

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
