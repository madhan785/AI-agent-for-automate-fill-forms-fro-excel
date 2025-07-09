import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from fuzzywuzzy import process

# === Load Excel values exactly as displayed (with formatting) ===
wb = load_workbook("Hhh.xlsx", data_only=False)
ws = wb.active

# Extract column headers
headers = [cell.value for cell in ws[1]]

# Build list of dictionaries for each row
data = []
for row in ws.iter_rows(min_row=2, values_only=False):
    row_data = {}
    for header, cell in zip(headers, row):
        row_data[header] = cell.value if cell.data_type != "n" else cell.number_format and cell.value and cell._value
    data.append(row_data)

print(f"Total submissions to make: {len(data)}")

# === Set up ChromeDriver ===
options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(), options=options)

# === Google Form Link ===
form_url = "https://forms.gle/7MEhV43oD2uqtnk1A"

# === Function to fill form ===
def fill_google_form(row_data):
    driver.get(form_url)
    time.sleep(3)

    questions = driver.find_elements(By.CSS_SELECTOR, "div[role='listitem']")

    for q in questions:
        try:
            q_text = q.text.split("\n")[0].strip()
            best_match, score = process.extractOne(q_text, headers)

            if score >= 70:
                answer = str(row_data[best_match]).strip()

                try:
                    input_box = q.find_element(By.CSS_SELECTOR, 'input[type="text"]')
                    input_box.send_keys(answer)
                except:
                    try:
                        text_area = q.find_element(By.CSS_SELECTOR, 'textarea')
                        text_area.send_keys(answer)
                    except:
                        print(f"❌ Couldn’t fill: {q_text}")
            else:
                print(f"⚠️ No good match for question: {q_text}")

        except Exception as e:
            print(f"⚠️ Error processing question: {e}")

    try:
        submit_button = driver.find_element(By.XPATH, "//span[contains(text(),'Submit')]/ancestor::div[@role='button']")
        submit_button.click()
        print("✅ Form submitted.")
        time.sleep(2)
    except Exception as e:
        print("❌ Submit button error:", e)

# === Loop through each row and fill form ===
for index, row in enumerate(data):
    print(f"\n🚀 Submitting row {index + 1}...")
    fill_google_form(row)
    if index < len(data) - 1:
        print("⏳ Waiting 10 seconds before next submission...")
        time.sleep(10)

driver.quit()
