# Word Template Auto-Filler V2.0 (GUI)

📄 A simple Python-based tool for quickly filling out Word templates with Excel data in just a few clicks.

## 🚀 Features

- ✅ Easy-to-use graphical interface (built with Tkinter)
- 📄 Select an Excel file and one or more Word templates via file picker
- 🔁 Automatically fills Word .docx templates using placeholders like <<name>>, <<birth_date>>, etc.
- 📊 Works with both text and table cells
- 📂 Output is saved in a separate (auto-created) folder to avoid overwriting
- 🖱️ One-click .exe version available – no Python installation needed
- 🔍 Clickable links to open generated Word files directly
- 🧭 Scrollable output list with file count and open folder button

## 📥 How to Use

1. Launch the app:
    - python gui.py (if running from source)
    - or double-click Word Template Auto-Filler.exe (no install needed)
2. Use the interface to:
    - Select your Excel .xlsx file
    - Select one or more Word .docx template files
    - Optionally select an output folder (or leave blank to auto-create)
    - Click ▶ **Fill Templates**
3. After generation:
    - Click on the file names to open generated Word files
    - Click the 📂 icon to open the output folder

## 🧩 Excel Format

   | key         | label            | value        |
   |-------------|------------------|--------------|
   | name        | Full Name        | Jane Doe     |
   | birth_date  | Date of Birth    | 1990-01-01   |
   | city        | City of Address  | Budapest     |


The program will replace << name >> in the Word templates with Jane Doe, and so on.

⚙️ 👉 [Download Auto-Filler-GUI](https://github.com/Csikito/mini_word_filler/releases/tag/v2.0)

## 🗺️ GUI Preview
<img width="952" height="382" alt="gui" src="https://github.com/user-attachments/assets/b0339c15-c812-4af5-a377-d838463c4ec1" />

## 📄 License

MIT – free to use, modify and distribute.

---

Created by Csikitó
