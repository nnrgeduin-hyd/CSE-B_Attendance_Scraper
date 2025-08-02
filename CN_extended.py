from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor, as_completed
from shutil import which
from datetime import datetime
import gspread
import time

# Constants
SHEET_ID = "1YFyoCRGsYWzZ4rxHyZt6KnWZLfjcrKiGyyrN9sytKeA"
SHEET_NAME = "CN"
ATTENDANCE_SHEET = "Attendence CSE-B(2023-27)"
SUBJECT = "CN"
MAX_THREADS = 15
MAX_ATTEMPTS = 3

# Setup credentials and Google Sheets client
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(credentials)

# Setup headless Chrome options
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = which("google-chrome") or which("chrome")
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--no-sandbox")

# Get worksheet and prepare columns
sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
rows = sheet.get_all_values()
timestamp = datetime.now().strftime("%d-%m %H:%M")
sheet.update_cell(10, len(rows[10]) + 1, timestamp)
col_index = len(rows[10]) + 1

# Get roll numbers from Attendence Sheet (B27:B91)
att_sheet = client.open_by_key(SHEET_ID).worksheet(ATTENDANCE_SHEET)
roll_cells = att_sheet.range("B27:B91")
roll_map = {cell.value.strip(): cell.row for cell in roll_cells if cell.value.strip()}
roll_numbers = list(roll_map.keys())
rollP_list = [r + "P" for r in roll_numbers]

# Function to scrape both attendance % and classes attended
def scrape_attendance(rollP):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(30)
            wait = WebDriverWait(driver, 5)

            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(rollP)
            driver.find_element(By.ID, "btnNext").click()
            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(rollP)
            driver.find_element(By.ID, "btnSubmit").click()
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))
            rows = driver.find_element(By.ID, "ctl00_cpStud_grdSubject").find_elements(By.TAG_NAME, "tr")[1:]

            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) < 6:
                    continue
                subject_name = cols[1].text.upper().strip()
                if SUBJECT in subject_name:
                    attendance_percent = cols[5].text.strip()
                    classes_attended = cols[4].text.strip()
                    if attendance_percent and attendance_percent != "&nbsp;":
                        return (rollP[:-1], attendance_percent, classes_attended)
            return (rollP[:-1], None, None)

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed for {rollP}: {e}")
        finally:
            try:
                driver.quit()
            except:
                pass
    print(f"❌ Failed to scrape {rollP}")
    return (rollP[:-1], None, None)

# Function to insert "Classes Attended" into Attendence CSE-B(2023-27) -> H27:H91
def insert_classes_attended(results):
    new_sheet = client.open_by_key(SHEET_ID).worksheet(ATTENDANCE_SHEET)
    roll_cells = new_sheet.range("B27:B91")
    roll_map = {cell.value.strip(): cell.row for cell in roll_cells if cell.value.strip()}

    class_cells = []
    for roll, _, classes in results:
        if not classes or roll not in roll_map:
            print(f"⚠️ No class data for {roll}")
            continue
        row = roll_map[roll]
        class_cells.append(gspread.Cell(row=row, col=8, value=classes))  # H = col 8
        print(f"📘 {roll} => {classes} classes attended")

    if class_cells:
        new_sheet.update_cells(class_cells)
        print(f"🟦 Inserted classes attended for {len(class_cells)} students.")

# Main driver
def main():
    results = []
    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = [executor.submit(scrape_attendance, rollP) for rollP in rollP_list]
        for future in as_completed(futures):
            result = future.result()
            if result:
                results.append(result)

    # Insert attendance % into CN sheet
    batch_cells = []
    for roll, attendance, _ in results:
        if not attendance:
            print(f"⚠️ No data for {roll}")
            continue
        if roll in roll_map:
            row = roll_map[roll]
            batch_cells.append(gspread.Cell(row=row, col=col_index, value=attendance + " %"))
            print(f"✅ {roll} => {attendance}%")
        else:
            print(f"❌ Roll not in sheet: {roll}")

    if batch_cells:
        sheet.update_cells(batch_cells)
        print(f"🟢 Inserted attendance % for {len(batch_cells)} students.")

    # Insert classes attended into main attendance sheet
    insert_classes_attended(results)

if __name__ == "__main__":
    main()
