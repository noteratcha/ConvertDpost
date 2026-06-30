---
name: dpost_converter
description: Skills for converting mailing labels in PDF files into DPost Excel import sheets and managing the Tkinter GUI.
---
# DPost PDF Converter Skill

Use this skill when modifying the PDF mailing label conversion logic or updating the Tkinter GUI for the DPost application.

## Core Architecture
- **Backend ([convert_dpost.py](file:///d:/Note/โปรเจค/Convert%20Dpost/convert_dpost.py))**:
  - Handles parsing logic (using `pypdf.PdfReader` and regex).
  - Normalizes Thai addresses (tambon, amphur, province, zipcode).
  - Converts parsed records to pandas DataFrames using [records_to_dataframe](file:///d:/Note/โปรเจค/Convert%20Dpost/convert_dpost.py#L348).
- **Frontend ([gui.py](file:///d:/Note/โปรเจค/Convert%20Dpost/gui.py))**:
  - Desktop UI built using `tkinter` and `ttk`.
  - Multi-threaded processing to keep the GUI responsive.
  - Redirects console prints directly to the GUI Log window.

## Custom Constraints
- Version format must always follow: `Year.MonthDay.HourMinute` (e.g. `2026.0630.1453`).
- Custom Tkinter buttons must use `disabledforeground` instead of `disabledbackground`.
