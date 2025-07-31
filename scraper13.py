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
MAX_THREADS = 1
BASE_PREFIX = "237Z1A05"

SUBJECT_SHEETS = [
    "Overall %", "CN", "DEVOPS", "PPL", "NLP", "DAA",
    "CN LAB", "DEVOPS LAB", "ACS LAB", "IPR",
    "Sports", "Men", "Association", "Library"
]

SUBJECT_ALIASES = {
    "CN": "CN",
    "DEVOPS": "DEVOPS",
    "PPL": "PPL",
    "NLP": "NLP",
    "DAA": "DAA",
    "CN LAB": "CN LAB",
    "DEVOPS LAB": "DEVOPS LAB",
    "ACS LAB": "ACS LAB",
    "IPR": "IPR",
    "SPORTS": "Sports",
    "MEN": "Men",
    "ASSOC": "Association",
    "ASSOCIATION": "Association",
    "LIB": "Library",
    "LIBRARY": "Library"
}

ATTENDED_RANGES = {
    "DAA": "F",
    "CN": "H",
    "DEVOPS": "J",
    "PPL": "L",
    "NLP": "N",
    "CN LAB": "P",
    "DEVOPS LAB": "R",
    "ACS LAB": "T",
    "IPR": "V",
    "Sports": "X",
    "Men": "Z",
    "Association": "AB",
    "Library": "AD"
}

CLEAR_RANGES = [
    "D8:D20", "J8:J20", "F27:F91", "H27:H91", "J27:J91", "L27:L91", "N27:N91",
    "P27:P91", "R27:R91", "T27:T91", "V27:V91", "X27:X91", "Z27:Z91", "AB27:AB91", "AD27:AD91"
]

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheets = {name: client.open_by_key(SHEET_ID).worksheet(name) for name in SUBJECT_SHEETS}

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

def generate_roll_numbers():
    rolls = []
    for num in range(72, 100):
        if str(num) in ["80", "88"]:
            continue
        rolls.append(BASE_PREFIX + str(num))
    for letter in ["A", "B", "C", "D"]:
        for d in range(0, 10):
            rolls.append(BASE_PREFIX + f"{letter}{d}")
    rolls.append("237Z1A05A8")
    return rolls

def prepare_new_column(sheet):
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, timestamp)
    return 3

def get_roll_row_mapping(sheet):
    all_rows = sheet.get_all_values()
    return {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}

def clear_attendance_cells():
    sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")
    for rng in CLEAR_RANGES:
        sheet.batch_clear([rng])
        time.sleep(1.5)

def process_roll(rollP):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(10)
            wait = WebDriverWait(driver, 5)

            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(rollP)
            driver.find_element(By.ID, "btnNext").click()
            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(rollP)
            driver.find_element(By.ID, "btnSubmit").click()
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_lblTotalPercentage")))
            overall_attendance = driver.find_element(By.ID, "ctl00_cpStud_lblTotalPercentage").text.strip()

            subjects_table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
            rows = subjects_table.find_elements(By.TAG_NAME, "tr")[1:]

            subject_data = {"Overall %": overall_attendance}
            attended_map = {}
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) < 6:
                    continue
                subject_text = cols[1].text.upper().split(":")[0].strip()
                attendance_percent = cols[5].text.strip()
                classes_attended = cols[3].text.strip()
                key = SUBJECT_ALIASES.get(subject_text)
                if key and attendance_percent and attendance_percent != "&nbsp;":
                    subject_data[key] = attendance_percent
                    if classes_attended and classes_attended != "&nbsp;":
                        attended_map[key] = classes_attended

            return (rollP[:-1], subject_data, attended_map)

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed for {rollP} — {e}")
            time.sleep(0.5)
        finally:
            try:
                driver.quit()
            except:
                pass

    print(f"❌ Failed to scrape {rollP}")
    return (rollP[:-1], {}, {})

def run_parallel_scraping():
    rolls = generate_roll_numbers()
    roll_with_p = [r + "P" for r in rolls]
    roll_to_row = {sheet: get_roll_row_mapping(sheets[sheet]) for sheet in SUBJECT_SHEETS}
    col_index = {sheet: prepare_new_column(sheets[sheet]) for sheet in SUBJECT_SHEETS}
    clear_attendance_cells()

    class_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")
    attended_data = {subject: {} for subject in ATTENDED_RANGES}

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {executor.submit(process_roll, roll): roll for roll in roll_with_p}
        for future in as_completed(futures):
            roll, subj_data, attended_map = future.result()
            if not subj_data:
                continue
            for subject, attendance in subj_data.items():
                if roll in roll_to_row.get(subject, {}):
                    row = roll_to_row[subject][roll]
                    col = col_index[subject]
                    value = attendance if subject == "Overall %" else attendance + " %"
                    try:
                        sheets[subject].update_cell(row, col, value)
                        print(f"✅ Updated {subject} → {roll}: {value}")
                        time.sleep(3.1)
                    except Exception as e:
                        print(f"❌ Error writing to {subject} for {roll}: {e}")
                else:
                    print(f"⚠️ Roll {roll} not found in sheet: {subject}")

            for subject, attended in attended_map.items():
                if subject in attended_data:
                    attended_data[subject][roll] = attended

    for subject, col_letter in ATTENDED_RANGES.items():
        if subject in attended_data:
            rolls_in_order = [r[:-1] for r in generate_roll_numbers()]
            values = [[attended_data[subject].get(roll, "")] for roll in rolls_in_order]
            try:
                cell_range = f"{col_letter}27:{col_letter}{26 + len(values)}"
                class_sheet.update(cell_range, values)
                print(f"✅ Inserted attended classes for {subject} into {cell_range}")
                time.sleep(2.5)
            except Exception as e:
                print(f"❌ Failed to update attended for {subject}: {e}")

if __name__ == "__main__":
    run_parallel_scraping()
