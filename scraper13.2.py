import os
import time
import gspread
from datetime import datetime
from zoneinfo import ZoneInfo
from shutil import which
from concurrent.futures import ThreadPoolExecutor, as_completed
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === CONFIG ===
SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"
BASE_PREFIX = "237Z1A05"
CREDENTIAL_FILES = [f"credentials{i}.json" for i in range(1, 15)]
CURRENT_CRED_INDEX = 0
MAX_ATTEMPTS = 3
MAX_THREADS = 8

SUBJECT_SHEETS = [
    "Overall %", "CN", "DEVOPS", "PPL", "NLP", "DAA",
    "CN LAB", "DEVOPS LAB", "ACS LAB", "IPR",
    "SPORTS", "NPTEL", "ASSOCIATION", "LIB/MEN"
]

SUBJECT_ALIASES = {
    "CN": "CN", "DEVOPS": "DEVOPS", "PPL": "PPL", "NLP": "NLP", "DAA": "DAA",
    "CN LAB": "CN LAB", "DEVOPS LAB": "DEVOPS LAB", "ACS LAB": "ACS LAB", "IPR": "IPR",
    "SPORTS": "SPORTS", "NPTEL": "NPTEL", "ASSOC": "ASSOCIATION", "ASSOCIATION": "ASSOCIATION",
    "LIB/MEN": "LIB/MEN", "LIB/MEN": "LIB/MEN"
}

SUBJECT_ClassesAttended_RANGES = {
    "DAA": "F27:F91",
    "CN": "H27:H91",
    "DEVOPS": "J27:J91",
    "PPL": "L27:L91",
    "NLP": "N27:N91",
    "CN LAB": "P27:P91",
    "DEVOPS LAB": "R27:R91",
    "ACS LAB": "T27:T91",
    "IPR": "V27:V91",
    "SPORTS": "X27:X91",
    "MENTORING": "Z27:Z91",
    "ASSOCIATION": "AB27:AB91",
    "LIBRARY": "AD27:AD91"
}

# === GOOGLE SHEETS SETUP ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

def get_gspread_client():
    cred_file = CREDENTIAL_FILES[CURRENT_CRED_INDEX]
    creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, scope)
    return gspread.authorize(creds)

def switch_credentials():
    global CURRENT_CRED_INDEX
    CURRENT_CRED_INDEX = (CURRENT_CRED_INDEX + 1) % len(CREDENTIAL_FILES)
    print(f"🔄 Switched to credential file: {CREDENTIAL_FILES[CURRENT_CRED_INDEX]}")
    return get_gspread_client()

def safe_call(func, *args, **kwargs):
    global client
    for _ in range(len(CREDENTIAL_FILES)):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429:
                print("⚠️ Rate limit. Switching creds...")
                client = switch_credentials()
                refresh_sheets()
            else:
                raise e
    raise RuntimeError("All credentials exhausted.")

def refresh_sheets():
    global sheets, class_sheet
    sheets = {name: client.open_by_key(SHEET_ID).worksheet(name) for name in SUBJECT_SHEETS}
    class_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

client = get_gspread_client()
refresh_sheets()

# === SELENIUM SETUP ===
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

# === UTILS ===
def generate_roll_numbers():
    rolls = [BASE_PREFIX + str(n) for n in range(72, 100) if str(n) not in ["80", "88"]]
    rolls += [BASE_PREFIX + f"{l}{d}" for l in "ABCD" for d in range(10)]
    return [r + "P" for r in rolls]

def get_roll_row_mapping(sheet):
    rows = safe_call(sheet.get_all_values)
    return {row[0].strip(): idx for idx, row in enumerate(rows[10:], start=11) if row and row[0].strip()}

def prepare_new_column(sheet):
    timestamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %I:%M %p")
    safe_call(sheet.insert_cols, [[]], 3)
    safe_call(sheet.update_cell, 10, 3, timestamp)
    return 3

def extract_classes_held(roll):
    try:
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 10)
        driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
        wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(roll)
        driver.find_element(By.ID, "btnNext").click()
        wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(roll)
        driver.find_element(By.ID, "btnSubmit").click()
        wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()
        wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))
        rows = driver.find_element(By.ID, "ctl00_cpStud_grdSubject").find_elements(By.TAG_NAME, "tr")[1:-1]
        held = [r.find_elements(By.TAG_NAME, "td")[3].text.strip() or "0" for r in rows if len(r.find_elements(By.TAG_NAME, "td")) >= 4]
        return held + ["0"] * (13 - len(held))
    except:
        return ["0"] * 13
    finally:
        try: driver.quit()
        except: pass

def process_roll(roll):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(10)
            wait = WebDriverWait(driver, 5)
            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(roll)
            driver.find_element(By.ID, "btnNext").click()
            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(roll)
            driver.find_element(By.ID, "btnSubmit").click()
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()
            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_lblTotalPercentage")))

            overall = driver.find_element(By.ID, "ctl00_cpStud_lblTotalPercentage").text.strip()
            table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
            rows = table.find_elements(By.TAG_NAME, "tr")[1:]

            percent_data = {"Overall %": overall}
            attended_data = {}

            for r in rows:
                cols = r.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 6:
                    subject = cols[1].text.upper().split(":")[0].strip()
                    percent = cols[5].text.strip()
                    attended = cols[4].text.strip()
                    key = SUBJECT_ALIASES.get(subject)
                    if key:
                        if percent:
                            percent_data[key] = percent
                        if attended:
                            attended_data[key] = attended

            return (roll[:-1], percent_data, attended_data)
        except:
            time.sleep(0.5)
        finally:
            try: driver.quit()
            except: pass
    return (roll[:-1], {}, {})

# === MAIN ===
def run_fast_scraper():
    global client
    rolls = generate_roll_numbers()
    roll_map = {s: get_roll_row_mapping(sheets[s]) for s in SUBJECT_SHEETS}
    col_index = {s: prepare_new_column(sheets[s]) for s in SUBJECT_SHEETS}
    batched_data = {s: [] for s in SUBJECT_SHEETS}
    attended_data_per_subject = {s: [] for s in SUBJECT_ClassesAttended_RANGES}

    print("⏳ Scraping class held...")
    class_72 = extract_classes_held(rolls[0])
    class_a8 = extract_classes_held("237Z1A05A8P")
    safe_call(class_sheet.update, "D8:D20", [[v] for v in class_72])
    safe_call(class_sheet.update, "J8:J20", [[v] for v in class_a8])
    print("✅ Inserted class held.")

    print("🧹 Clearing old classes attended data...")
    for subject, cell_range in SUBJECT_ClassesAttended_RANGES.items():
        safe_call(class_sheet.update, cell_range, [[""] for _ in range(65)])

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {executor.submit(process_roll, r): r for r in rolls}
        for f in as_completed(futures):
            roll, percent_data, attended_data = f.result()
            for subject, val in percent_data.items():
                if roll in roll_map.get(subject, {}):
                    row = roll_map[subject][roll]
                    col = col_index[subject]
                    batched_data[subject].append((row, col, val if subject == "Overall %" else val + " %"))

            for subject, val in attended_data.items():
                if subject in SUBJECT_ClassesAttended_RANGES:
                    if roll in roll_map.get(subject, {}):
                        row = roll_map[subject][roll]
                        attended_data_per_subject[subject].append((row, val))

    print("📝 Writing scraped percentage data...")
    for subject, updates in batched_data.items():
        if not updates:
            continue
        rows = list(set(row for row, col, val in updates))
        min_row, max_row = min(rows), max(rows)
        col = col_index[subject]
        cell_range = f"{chr(64 + col)}{min_row}:{chr(64 + col)}{max_row}"
        cell_objs = safe_call(sheets[subject].range, cell_range)
        cell_map = {(cell.row, cell.col): cell for cell in cell_objs}
        for row, col, val in updates:
            if (row, col) in cell_map:
                cell_map[(row, col)].value = val
        safe_call(sheets[subject].update_cells, list(cell_map.values()))
        print(f"✅ {subject} updated")

    print("📝 Writing classes attended data...")
    for subject, updates in attended_data_per_subject.items():
        if not updates:
            continue
        updates.sort()
        cell_range = SUBJECT_ClassesAttended_RANGES[subject]
        values = [[v] for _, v in updates]
        safe_call(class_sheet.update, cell_range, values)
        print(f"✅ {subject} classes attended updated")

if __name__ == "__main__":
    run_fast_scraper()
