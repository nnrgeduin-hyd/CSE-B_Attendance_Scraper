# ✅ Attendance Scraper with Auto-Rotating Google Credentials

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

SUBJECT_SHEETS = [
    "Overall %", "CN", "DEVOPS", "PPL", "NLP", "DAA",
    "CN LAB", "DEVOPS LAB", "ACS LAB", "IPR",
    "SPORTS", "MENTORING", "ASSOCIATION", "LIBRARY"
]

SUBJECT_ALIASES = {
    "CN": "CN", "DEVOPS": "DEVOPS", "PPL": "PPL", "NLP": "NLP", "DAA": "DAA",
    "CN LAB": "CN LAB", "DEVOPS LAB": "DEVOPS LAB", "ACS LAB": "ACS LAB", "IPR": "IPR",
    "SPORTS": "SPORTS", "MEN": "MENTORING", "ASSOC": "ASSOCIATION", "ASSOCIATION": "ASSOCIATION",
    "LIB": "LIBRARY", "LIBRARY": "LIBRARY"
}

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

def get_gspread_client():
    cred_file = CREDENTIAL_FILES[CURRENT_CRED_INDEX]
    if not os.path.exists(cred_file):
        raise FileNotFoundError(f"Credential file {cred_file} not found")
    creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, scope)
    return gspread.authorize(creds)

def switch_credentials():
    global CURRENT_CRED_INDEX
    CURRENT_CRED_INDEX = (CURRENT_CRED_INDEX + 1) % len(CREDENTIAL_FILES)
    print(f"\U0001f501 Switched to credential file: {CREDENTIAL_FILES[CURRENT_CRED_INDEX]}")
    return get_gspread_client()

def safe_call(func, *args, **kwargs):
    global client, sheets, class_sheet
    for _ in range(len(CREDENTIAL_FILES)):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429:
                print("\u26a0\ufe0f Rate limit hit. Switching credentials...")
                client = switch_credentials()
                sheets = {name: client.open_by_key(SHEET_ID).worksheet(name) for name in SUBJECT_SHEETS}
                class_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")
            else:
                raise e
    raise RuntimeError("\u274c All credentials hit the rate limit.")

# === INITIALIZE CLIENTS ===
client = get_gspread_client()
sheets = {name: client.open_by_key(SHEET_ID).worksheet(name) for name in SUBJECT_SHEETS}
class_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

def clear_attendance_sheet():
    ranges = [
        "D8:D20", "J8:J20", "F27:F91", "H27:H91", "J27:J91", "L27:L91",
        "N27:N91", "P27:P91", "R27:R91", "T27:T91", "V27:V91", "X27:X91",
        "Z27:Z91", "AB27:AB91", "AD27:AD91"
    ]
    for rng in ranges:
        safe_call(class_sheet.batch_clear, [rng])
    print("\U0001f9f9 Cleared Attendence CSE-B(2023-27) ranges.")

def generate_roll_numbers():
    rolls = [BASE_PREFIX + str(n) for n in range(72, 100) if str(n) not in ["80", "88"]]
    rolls += [BASE_PREFIX + f"{l}{d}" for l in "ABCD" for d in range(10)]
    return rolls

def prepare_new_column(sheet):
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    safe_call(sheet.insert_cols, [[]], 3)
    safe_call(sheet.update_cell, 10, 3, timestamp)
    return 3

def get_roll_row_mapping(sheet):
    all_rows = safe_call(sheet.get_all_values)
    return {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}

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
        held = [r.find_elements(By.TAG_NAME, "td")[3].text.strip() or "0" for r in rows if len(r.find_elements(By.TAG_NAME, "td")) >= 4]
        return held + ["0"] * (13 - len(held))
    except Exception as e:
        print(f"\u274c Error fetching classes held for {rollP}: {e}")
        return ["0"] * 13
    finally:
        try: driver.quit()
        except: pass

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
                if len(cols) < 6: continue
                subject = cols[1].text.upper().split(":")[0].strip()
                percent = cols[5].text.strip()
                key = SUBJECT_ALIASES.get(subject)
                if key and percent and percent != "&nbsp;":
                    data[key] = percent
            return (rollP[:-1], data)
        except Exception as e:
            print(f"\u26a0\ufe0f Attempt {attempt} failed for {rollP} — {e}")
            time.sleep(0.5)
        finally:
            try: driver.quit()
            except: pass
    print(f"\u274c Failed to scrape {rollP}")
    return (rollP[:-1], {})

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

def run_parallel_scraping():
    global client, sheets, class_sheet
    clear_attendance_sheet()
    rolls = generate_roll_numbers()
    roll_with_p = [r + "P" for r in rolls]
    roll_to_row = {s: get_roll_row_mapping(sheets[s]) for s in SUBJECT_SHEETS}
    col_index = {s: prepare_new_column(sheets[s]) for s in SUBJECT_SHEETS}

    held_72 = extract_classes_held(roll_with_p[0])
    safe_call(class_sheet.update, "D8:D20", [[v] for v in held_72])
    print("\u2705 Inserted Classes Held from 72 into D8:D20")

    held_a8 = extract_classes_held("237Z1A05A8P")
    safe_call(class_sheet.update, "J8:J20", [[v] for v in held_a8])
    print("\u2705 Inserted Classes Held from A8 into J8:J20")

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
                    try:
                        safe_call(sheets[subject].update_cell, row, col, val)
                        print(f"\u2705 Updated {subject} → {roll}: {val}")
                    except Exception as e:
                        print(f"\u274c Error writing to {subject} for {roll}: {e}")
                else:
                    print(f"\u26a0\ufe0f Roll {roll} not found in sheet: {subject}")

if __name__ == "__main__":
    run_parallel_scraping()
