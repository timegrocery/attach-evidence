# 📎 Attach Evidence Links

A small Python tool that helps turn codes like `[1.2-001]` in a Word document into clickable links,  
based on a mapping in an Excel file.  
Mainly made to save time instead of adding links one by one.

---

## ✨ Features

- 🔍 Find codes in a Word `.docx` that match a pattern (e.g., `[1.2-001]`).
- 📑 Read an Excel file to match each code with a link (and an optional tooltip).
- 🔗 Add hyperlinks in place of the codes, keeping other text intact.
- 🗂 Works for both normal paragraphs and inside tables.

---

## 🛠 How it works

1. **Regex pattern**:  
   Pattern used: ``\[(?P<code>\d+\.\d+-\d{3})\]``  
2. **Excel reading**: Uses the columns you select for:
   - Code (required)
   - Link (required)
   - Summary/tooltip (optional)
3. **Word editing**: Inserts clickable links with blue underline style.
4. **Table support**: Checks every cell in every table.

---

## 📦 Requirements

- Python 3.9+
- [`python-docx`](https://pypi.org/project/python-docx/)
- [`openpyxl`](https://pypi.org/project/openpyxl/)
- [`ttkbootstrap`](https://pypi.org/project/ttkbootstrap/) (for the GUI)