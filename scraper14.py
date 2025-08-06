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
import os
import json

# === CONFIG ===
SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"
MAX_ATTEMPTS = 3
MAX_THREADS = 10
BASE_PREFIX = "237Z1A05"
CREDENTIAL_FILES = [f"credentials{i}.json" for i in range(1, 15)]
CURRENT_CRED_INDEX = 0
RATE_LIMIT_WAIT = 60  # Seconds to wait before retrying all credentials

SUBJECT_SHEETS = [
    "Overall %", "CN", "DEVOPS", "PPL", "NLP", "DAA",
    "CN LAB", "DEVOPS LAB", "ACS LAB", "IPR",
    "SPORTS", "MENTORING", "ASSOCIATION", "LIBRARY"
]

# === SUBJECT NAME NORMALIZATION ===
SUBJECT_ALIASES = {
    "CN": "CN", "DEVOPS": "DEVOPS", "PPL": "PPL", "NLP": "NLP", "DAA": "DAA",
    "CN LAB": "CN LAB", "DEVOPS LAB": "DEVOPS LAB", "ACS LAB": "ACS LAB", "IPR": "IPR",
    "SPORTS": "SPORTS", "MENTORING": "MENTORING", "ASSOCIATION": "ASSOCIATION", "ASSOCIATION": "ASSOCIATION",
    "LIBRARY": "LIBRARY", "LIBRARY": "LIBRARY"
}

# === SETUP GOOGLE SHEETS ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

def get_gspread_client(cred_index):
    cred_file = CREDENTIAL_FILES[cred_index]
    if not os.path.exists(cred_file):
        raise FileNotFoundError(f"Credential file {cred_file} not found")
    creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, scope)
    return gspread.authorize(creds)

def switch_credentials(current_index):
    return (current_index + 1) % len(CREDENTIAL_FILES)  # Loop back to 0 after last credential

# Initialize first client
client = get_gspread_client(CURRENT_CRED_INDEX)
sheets = {name: client.open_by_key(SHEET_ID).worksheet(name) for name in SUBJECT_SHEETS}
class_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

# === RETRY WITH ALL CREDENTIALS ===
def retry_with_all_credentials(operation, *args, operation_name="operation"):
    global client, sheets, class_sheet
    cred_index = CURRENT_CRED_INDEX
    attempted_all = False

    while True:
        try:
            client = get_gspread_client(cred_index)
            sheets = {name: client.open_by_key(SHEET_ID).worksheet(name) for name in SUBJECT_SHEETS}
            class_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")
            return operation(*args)  # Execute the operation
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429:
                print(f"⚠️ Rate limit hit for {CREDENTIAL_FILES[cred_index]} during {operation_name}")
                cred_index = switch_credentials(cred_index)
                if cred_index == CURRENT_CRED_INDEX and attempted_all:
                    print(f"⏳ All credentials rate-limited, waiting {RATE_LIMIT_WAIT}s before retrying...")
                    time.sleep(RATE_LIMIT_WAIT)
                    attempted_all = False
                else:
                    attempted_all = True
                    print(f"🔄 Switching to {CREDENTIAL_FILES[cred_index]}")
            else:
                raise e  # Re-raise non-rate-limit errors

# === CLEAR RANGES IN CLASS SHEET ===
def clear_attendance_sheet():
    def do_clear():
        ranges = [
            "D8:D20", "J8:J20", "F27:F91", "H27:H91", "J27:J91", "L27:L91",
            "N27:N91", "P27:P91", "R27:R91", "T27:T91", "V27:V91", "X27:X91",
            "Z27:Z91", "AB27:AB91", "AD27:AD91"
        ]
        for rng in ranges:
            class_sheet.batch_clear([rng])
        print("🧹 Cleared Attendence CSE-B(2023-27) ranges.")
    return retry_with_all_credentials(do_clear, operation_name="clear_attendance_sheet")

# === ADD COLUMN ===
def prepare_new_column(sheet):
    def do_prepare():
        ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
        timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
        sheet.insert_cols([[]], 3)
        sheet.update_cell(10, 3, timestamp)
        return 3
    return retry_with_all_credentials(do_prepare, operation_name=f"prepare_new_column for {sheet.title}")

# === ROLL MAPPING ===
def get_roll_row_mapping(sheet):
    def do_mapping():
        all_rows = sheet.get_all_values()
        return {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}
    return retry_with_all_credentials(do_mapping, operation_name=f"get_roll_row_mapping for {sheet.title}")

# === CLASSES HELD FOR ONE ROLL ===
def extract_classes_held(rollP):
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 10)
    try:
        driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
        wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(rollP)
        driver.find_element(By.ID, "btnNext").click()
        wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(rollP)
        driver.find_element(By.ID, "btnSubmit").click()
        wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()
        wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))

        rows = driver.find_element(By.ID, "ctl00_cpStud_grdSubject").find_elements(By.TAG_NAME, "tr")[1:-1]
        held = []
        for r in rows:
            cols = r.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 4:
                held.append(cols[3].text.strip() or "0")
        return held + ["0"] * (13 - len(held))
    except Exception as e:
        print(f"❌ Error fetching classes held for {rollP}: {e}")
        return ["0"] * 13
    finally:
        try:
            driver.quit()
        except:
            pass

# === SCRAPE ONE ROLL ===
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

            overall = driver.find_element(By.ID, "ctl00_cpStud_lblTotalPercentage").text.strip()
            table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
            rows = table.find_elements(By.TAG_NAME, "tr")[1:]
            data = {"Overall %": overall}
            for r in rows:
                cols = r.find_elements(By.TAG_NAME, "td")
                if len(cols) < 6:
                    continue
                subject = cols[1].text.upper().split(":")[0].strip()
                percent = cols[5].text.strip()
                key = SUBJECT_ALIASES.get(subject)
                if key and percent and percent != "&nbsp;":
                    data[key] = percent
            return (rollP[:-1], data)
        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed for {rollP} — {e}")
            time.sleep(0.5)
        finally:
            try:
                driver.quit()
            except:
                pass
    print(f"❌ Failed to scrape {rollP}")
    return (rollP[:-1], {})

# === CHROME OPTIONS ===
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

# === MAIN ===
def run_parallel_scraping():
    global client, sheets, class_sheet
    clear_attendance_sheet()
    rolls = generate_roll_numbers()
    roll_with_p = [r + "P" for r in rolls]
    roll_to_row = {s: get_roll_row_mapping(sheets[s]) for s in SUBJECT_SHEETS}
    col_index = {s: prepare_new_column(sheets[s]) for s in SUBJECT_SHEETS}

    # → Classes Held for 1st roll
    held_72 = extract_classes_held(roll_with_p[0])
    def update_d8():
        class_sheet.update("D8:D20", [[v] for v in held_72])
        print("✅ Inserted Classes Held from 72 into D8:D20")
    retry_with_all_credentials(update_d8, operation_name="update D8:D20")

    # → Classes Held for 237Z1A05A8
    held_a8 = extract_classes_held("237Z1A05A8P")
    def update_j8():
        class_sheet.update("J8:J20", [[v] for v in held_a8])
        print("✅ Inserted Classes Held from A8 into J8:J20")
    retry_with_all_credentials(update_j8, operation_name="update J8:J20")

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {executor.submit(process_roll, r): r for r in roll_with_p}
        for f in as_completed(futures):
            roll, data = f.result()
            if not data:
                continue
            for subject, val in data.items():
                if roll in roll_to_row.get(subject, {}):
                    row = roll_to_row[subject][roll]
                    col = col_index[subject]
                    val = val if subject == "Overall %" else val + " %"
                    def update_cell():
                        sheets[subject].update_cell(row, col, val)
                        print(f"✅ Updated {subject} → {roll}: {val}")
                    retry_with_all_credentials(update_cell, operation_name=f"update_cell {subject} for {roll}")
                else:
                    print(f"⚠️ Roll {roll} not found in sheet: {subject}")

if __name__ == "__main__":
    run_parallel_scraping()