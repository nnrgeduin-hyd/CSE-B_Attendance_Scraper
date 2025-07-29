from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor, as_completed
from shutil import which
from datetime import datetime
from zoneinfo import ZoneInfo
import gspread
import time

# === CONFIG ===
SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"
MAX_ATTEMPTS = 3
MAX_THREADS = 10
BASE_PREFIX = "237Z1A05"
SUBJECT_CODES = ["DAA", "CN", "DEVOPS", "PPL", "NLP"]

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheets = {subj: client.open_by_key(SHEET_ID).worksheet(subj) for subj in SUBJECT_CODES}

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

# === Generate Roll Numbers ===
def generate_roll_numbers():
    rolls = []
    for num in range(72, 100):
        if str(num) in ["80", "88"]: continue
        rolls.append(BASE_PREFIX + str(num))
    for letter in ["A", "B", "C", "D"]:
        for d in range(1, 10):
            rolls.append(BASE_PREFIX + f"{letter}{d}")
    return rolls

# === Insert Timestamp Column for Each Sheet ===
def prepare_column_in_sheets():
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    col_position_map = {}
    for subj, sheet in sheets.items():
        sheet.insert_cols([[]], 3)
        sheet.update_cell(10, 3, timestamp)
        col_position_map[subj] = 3
        print(f"📅 [{subj}] Inserted column C with timestamp {timestamp}")
    return col_position_map

# === Get Roll → Row Mapping ===
def get_roll_row_mapping(sheet):
    data = sheet.get_all_values()
    return {row[0].strip(): idx for idx, row in enumerate(data[10:], start=11) if row and row[0].strip()}

# === Scrape Attendance for One Roll ===
def process_roll(rollP):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 10)
            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")

            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(rollP)
            driver.find_element(By.ID, "btnNext").click()
            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(rollP)
            driver.find_element(By.ID, "btnSubmit").click()

            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()
            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))

            table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
            rows = table.find_elements(By.TAG_NAME, "tr")[1:]

            subj_attendance = {}
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 6:
                    raw_name = cols[1].text.strip()
                    percentage = cols[5].text.strip()
                    if ":" in raw_name:
                        short_name = raw_name.split(":")[0].strip()
                        if short_name in SUBJECT_CODES:
                            subj_attendance[short_name] = percentage if percentage else "0"

            print(f"✅ {rollP} → {subj_attendance}")
            return (rollP[:-1], subj_attendance)
        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed for {rollP}: {e}")
            time.sleep(1)
        finally:
            try: driver.quit()
            except: pass

    print(f"❌ Max attempts failed for {rollP}")
    return (rollP[:-1], {})

# === Main Scraping Runner ===
def run_scraper():
    rolls = generate_roll_numbers()
    rolls_with_P = [r + "P" for r in rolls]

    col_map = prepare_column_in_sheets()
    row_maps = {subj: get_roll_row_mapping(sheets[subj]) for subj in SUBJECT_CODES}

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {executor.submit(process_roll, roll): roll for roll in rolls_with_P}
        for future in as_completed(futures):
            roll, data = future.result()
            for subj, value in data.items():
                if roll in row_maps[subj]:
                    sheets[subj].update_cell(row_maps[subj][roll], col_map[subj], value)
                else:
                    print(f"⚠️ Roll {roll} not found in sheet {subj}")

    print("🎉 All subject-wise attendance updated!")

# === MAIN ===
if __name__ == "__main__":
    run_scraper()
