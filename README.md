# 🎓 RGPV Result Fetcher 🔍

A Python automation tool to fetch and store RGPV B.Tech student results in an Excel sheet using **Selenium**, **OCR (Tesseract)**, and **BeautifulSoup**.

---

## 👨‍💻 Designed By: [Moh Technology](https://www.youtube.com/@mohtechnology)

---

## 📦 Features

- Automatically fills RGPV result form
- Captures and decodes CAPTCHA using OCR
- Extracts student grades and stores in an Excel file
- Skips invalid entries with blank or zero-filled rows
- Output formatted cleanly in `.xlsx`

---

## 🛠️ Installation Guide

### 1. 📥 Clone or Download
```bash
git clone https://github.com/mohtechnology/RGPV-Result.git
cd rgpv-result-fetcher
````

### 2. 📦 Install Python Requirements

```bash
pip install -r requirements.txt
```

**requirements.txt**

```
selenium
pillow
pytesseract
requests
openpyxl
```

### 3. 📷 Install Tesseract OCR

#### For Windows

* Download from: [https://github.com/tesseract-ocr/tesseract/wiki](https://github.com/tesseract-ocr/tesseract/wiki)
* Install and note the path (e.g., `C:\Program Files\Tesseract-OCR\tesseract.exe`)
* Add the path in your code:

```python
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

#### For Linux

```bash
sudo apt update
sudo apt install tesseract-ocr
```

### 4. 🌐 Install Chrome WebDriver (Optional)

* Download: [https://chromedriver.chromium.org/downloads](https://chromedriver.chromium.org/downloads)
* Match the version with your Chrome browser.
* Add it to your system PATH or specify the path in the code:

```python
driver = webdriver.Chrome(executable_path='path/to/chromedriver')
```

---

## 🚀 How to Use

1. Add enrollment numbers in the script.
2. Run the script:

```bash
python new.py
```

3. Results will be saved in `results.xlsx`.

If result not found or CAPTCHA fails, the script adds a row with `0` in all grade columns.

---

## 📁 Output Format

The Excel sheet will look like:

| S.No | Enrollment | Name         | Semester | BT101 | BT102 | ... | SGPA | CGPA | Result              |
| ---- | ---------- | ------------ | -------- | ----- | ----- | --- | ---- | ---- | ------------------- |
| 1    | 0805CS2410 | XYZ          | 1        | C     | C     | ... | 5.57 | 5.57 | Fail in BT104,BT105 |
| 2    | 0805CS2411 |              |          | 0     | 0     | ... | 0    | 0    |                     |

---

## 🧠 Notes

* Run using a stable internet connection.
* Make sure your browser & WebDriver versions match.
* Use real enrollment numbers for accurate testing.

---

## 📜 License

MIT License

---

## 📞 Contact

For support or queries, visit [Moh Technology](https://www.youtube.com/@mohtechnology)
