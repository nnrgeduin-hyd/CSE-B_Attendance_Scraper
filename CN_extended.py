from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor, as_completed
from shutil import which
from datetime import datetime
import gspread
from zoneinfo import ZoneInfo

# === CONFIG ===
SUBJECT = "CN"
SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"
CREDENTIAL_FILE = "credentials3.json"
MAX_ATTEMPTS = 3
BASE_PREFIX = "237Z1A05"
THREADS = 15
BATCH_SIZE = THREADS

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIAL_FILE, scope)
client = gspread.authorize(creds)
subject_sheet = client.open_by_key(SHEET_ID).worksheet(SUBJECT)
main_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

# === Chrome Options ===
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--window-size=1280,800")

chrome_path = which("chromium-browser")
if chrome_path:
    chrome_options.binary_location = chrome_path

# === Helpers ===
def generate_roll_numbers():
    rolls = []
    for num in range(72, 100):
        if str(num) in ["80", "88"]:
            continue
        rolls.append(BASE_PREFIX + str(num))
    for letter in ["A", "B", "C", "D"]:
        for d in range(0, 10):
            rolls.append(BASE_PREFIX + f"{letter}{d}")
    return rolls

def get_roll_row_mapping_main_sheet():
    rows = main_sheet.get("B27:B91")
    return {row[0].strip(): idx + 27 for idx, row in enumerate(rows) if row and row[0].strip()}

def get_roll_row_mapping_subject_sheet():
    all_rows = subject_sheet.get_all_values()
    return {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}

def prepare_subject_column():
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    subject_sheet.insert_cols([[]], 3)
    subject_sheet.update_cell(10, 3, timestamp)
    return 3

def scrape_attendance(rollP):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(50)
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
                    percentage = cols[5].text.strip()
                    attended = cols[4].text.strip()
                    return (rollP[:-1], percentage, attended)

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

# === Main Execution ===
def main():
    main_map = get_roll_row_mapping_main_sheet()
    subj_map = get_roll_row_mapping_subject_sheet()
    col_index = prepare_subject_column()
    all_rolls = generate_roll_numbers()

    # Clear H27:H91
    main_sheet.batch_clear(["H27:H91"])

    for i in range(0, len(all_rolls), BATCH_SIZE):
        batch = all_rolls[i:i + BATCH_SIZE]
        results = []

        with ThreadPoolExecutor(max_workers=THREADS) as executor:
            futures = [executor.submit(scrape_attendance, roll + "P") for roll in batch]
            for future in as_completed(futures):
                result = future.result()
                if result:
                    results.append(result)

        subject_cells = []
        main_cells = []

        for roll, percent, attended in results:
            if percent and roll in subj_map:
                row = subj_map[roll]
                subject_cells.append(gspread.Cell(row=row, col=col_index, value=percent + " %"))
                print(f"✅ {roll} => {percent}%")

            if attended and roll in main_map:
                row = main_map[roll]
                main_cells.append(gspread.Cell(row=row, col=8, value=attended))  # col 8 = H
            else:
                print(f"❌ Missing roll {roll} in Attendence CSE-B(2023-27)")

        if subject_cells:
            subject_sheet.update_cells(subject_cells)
            print(f"🟢 Subject sheet: Inserted {len(subject_cells)} % cells")

        if main_cells:
            main_sheet.update_cells(main_cells)
            print(f"🟢 Main sheet: Inserted {len(main_cells)} attended values")

if __name__ == "__main__":
    main()
