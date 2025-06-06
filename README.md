# 📄 RGPV Result Fetcher and Excel Exporter

**Automatically fetch RGPV student results** using enrollment number, decode CAPTCHA via OCR, extract subject grades, and **store everything into a neat Excel sheet** — all from a Python script.

> Designed By **Moh Technology**

---

## 🚀 Features

- 🔎 Fetch results from RGPV portal directly.
- 🤖 Solve CAPTCHA using Tesseract OCR.
- 📋 Extract all student information including:
  - Name, Roll Number, Program, Branch, Semester
  - Subject-wise Grades
  - SGPA and CGPA
- 📁 Store data row-by-row in an Excel sheet.
- 🧠 Automatically handles CAPTCHA failure or result not found.
- 📐 Column widths are auto-managed based on content.
- 🟨 Shows `0` in every column if CAPTCHA fails or result not found.

---

## 🧰 Requirements

### Python Libraries

Install all required Python libraries using pip:

```bash
pip install requests beautifulsoup4 pillow openpyxl
