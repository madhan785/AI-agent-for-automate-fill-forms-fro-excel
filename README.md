Here's a clean, professional README.md file you can use for your GitHub project:

📝 README.md

# 🧠 Google Form Auto-Filler using Python + Excel + Selenium

This project automates Google Form submissions using data from an Excel file. It uses Python, Selenium, and openpyxl to fill each form entry exactly as it appears in the spreadsheet — including correctly formatted values like percentages (e.g., 30%), IDs, and long strings.

## 🔧 Features

* ✅ Automatically fills and submits Google Forms using Excel data
* ✅ Supports text fields and multiline responses
* ✅ Preserves formatting (like 29.00%, 0.15, etc.)
* ✅ Handles row-by-row submissions with delay
* ✅ Uses fuzzy matching to map Excel column names to form questions
* ✅ Works with any Google Form (no need to modify the form)

## 📁 Folder Structure

```
.
├── auto_google_form.py        # Main Python script
├── Hhh.xlsx                   # Input Excel file with form data
├── chromedriver.exe           # Chrome WebDriver (matching your Chrome version)
├── README.md                  # Project overview and instructions
```

## 📦 Requirements

Install required Python packages:

```bash
pip install selenium openpyxl fuzzywuzzy
```

Make sure you have:

* Python 3.x installed
* Google Chrome installed
* ChromeDriver (matching your Chrome version) placed in the same folder

## 📄 Input Format

Hhh.xlsx should contain a header row with question labels and one row per form submission. Example:

| Name     | Email                                       | Score | Percent | ID     |
| -------- | ------------------------------------------- | ----- | ------- | ------ |
| John Doe | [john@example.com](mailto:john@example.com) | 85    | 29.00%  | 123456 |
| Jane Doe | [jane@example.com](mailto:jane@example.com) | 72    | 0.15    | 987654 |

Note: All cell formatting (like % or float) is preserved in the final form submission.

## 🚀 How to Run

Open your terminal and run:

```bash
python auto_google_form.py
```

You can also run it from a specific Python version like this:

```bash
"C:\Path\To\Your\Python\python.exe" auto_google_form.py
```

The script will:

* Launch Chrome
* Open the target Google Form
* Submit each row from Excel as a separate response
* Wait 10 seconds between submissions (configurable)

## 🧪 Test It With Your Own Form

Replace the form\_url variable in the script with your own Google Form link:

```python
form_url = "https://forms.gle/YOUR_FORM_LINK"
```

Make sure the form accepts responses and matches the questions in your Excel sheet.

## 🛠 Customization

You can modify the script to:

* Submit only selected rows
* Log responses to a .txt file
* Run headless (no Chrome window)
* Add error handling or screenshots

