# Excel Row Matcher

A GUI-based Excel analysis tool that filters rows based on specific conditions using Python, Tkinter, and OpenPyXL. This tool is designed for quick data validation and filtering without requiring manual Excel work.

---

## 🚀 Features

* Simple and user-friendly GUI (Tkinter)
* Automatically loads and scans Excel files
* Filters rows based on custom logic:

  * Column **D** must contain: `WS-C2960-24TT-L`
  * Column **C** must be greater than or equal to **20**
  * Column **B** must **not** contain: `atm`, `bje`
* Displays all matching rows in a scrollable text window
* Useful for network audits, data cleaning, and bulk validation

---

## 📦 Project Structure

```
excel-row-matcher/
│
├── scripts/
│   └── excel_row_matcher.py   # Main GUI filtering script
│
├── samples/
│   └── sample.xlsx            # Example Excel file (optional)
│
├── requirements.txt           # Dependencies
│
└── README.md                  # Documentation
```

---

## 🛠 Requirements

* Python 3.8+
* Tkinter (comes bundled with most Python installations)
* OpenPyXL

Install dependencies:

```
pip install -r requirements.txt
```

---

## ⚙️ Usage

### 1. Run the application

```
python scripts/excel_row_matcher.py
```

### 2. Click the button **"Open Excel File"**

Select the `.xlsx` file you want to analyze.

### 3. The application will:

* Load the file
* Scan all rows
* Apply filtering logic
* Show matching entries in the scrollable text box

---

## 🧠 Filtering Rules

The following conditions must be **all true** for a row to be considered a match:

### ✔ Column D contains:

```
WS-C2960-24TT-L
```

### ✔ Column C is a number and:

```
>= 20
```

### ✔ Column B does NOT contain (case-insensitive):

```
atm
bje
```

---

## 🧪 Example Output

```
Matching Rows:
--------------------------------------------------
Row 16: B=Port-11, C=24, D=WS-C2960-24TT-L
Row 22: B=USER-45, C=31, D=WS-C2960-24TT-L
--------------------------------------------------
Total Matching Rows: 2
```

---

## 🧰 Notes

* The script uses the **active sheet** of the Excel file
* Assumes data starts from **row 2** (row 1 = header)
* Non-numeric values in column C are treated as zero

---

## 🔐 Security and Safety

* No internet access required
* No data is uploaded anywhere
* All processing happens locally

---

## 🧩 Possible Future Enhancements

* Export matching rows to a new Excel file
* Add color-coded results
* Support for selecting different sheets
* Allow custom filtering rules
* Create EXE version using PyInstaller
* Add log file generation

---

## 🤝 Contributions

Pull requests are welcome.

If you find this project useful, please give it a ⭐ on GitHub!
