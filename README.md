# Word Template Auto-Filler

📄 A simple Python-based tool for quickly filling out Word templates with Excel data in just a few clicks.

## 🚀 Features

- Fill multiple `.docx` template files with data from one Excel file
- Use custom placeholders like `<<name>>` or `<<birth_date>>` in your Word templates
- Supports both paragraphs and table cells
- Outputs each generated document into a separate folder (safe from overwriting)
- No installation needed – one-click `.exe` version available (Windows)

## 📁 Folder Structure

mini_word_filler/

    ├── main.py # Main script
    ├── input.xlsx # Excel file with key-value pairs
    ├── templates/ # Folder for .docx template files
    ├── output/ # Folder for generated files


## 📥 How to Use

1. Place your Word `.docx` template files into the `templates/` folder.
2. Create an `input.xlsx` file with:
   - Column A: template keys (e.g. `name`, `birth_date`)
   - Column B: labels (optional, just for clarity)
   - Column C: actual values to fill in

   Example:

   | key         | label            | value        |
   |-------------|------------------|--------------|
   | name        | Full Name        | Jane Doe     |
   | birth_date  | Date of Birth    | 1990-01-01   |
   | city        | City of Address  | Budapest     |

3. Run the program:
   - If using the `.py`: `python main.py`
   - If using `.exe`: just double-click it
4. Output will be saved inside `output/` (with time- or versioned subfolders)


⚙️ 👉 [Download Windows executable (.exe)](https://github.com/Csikito/mini_word_filler/releases/tag/v1.0)

## 🗺️ Guide - GIF

![filler_v1](https://github.com/user-attachments/assets/9b40421c-36b2-4125-8f22-fcc1980b814d)

## 📄 License

MIT – free to use, modify and distribute.

---

Created by Csikitó
