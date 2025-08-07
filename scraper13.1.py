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

# === CONFIG ===
SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"
MAX_ATTEMPTS = 3
MAX_THREADS = 10
BASE_PREFIX = "237Z1A05"
CREDENTIAL_FILES = [f"credentials{i}.json" for i in range(1, 15)]
SUBJECT_SHEETS = [
    "Overall %", "CN", "DEVOPS", "PPL", "NLP", "DAA",
    "CN LAB", "DEVOPS LAB", "ACS LAB", "IPR",
    "SPORTS", "MENTORING", "ASSOCIATION", "LIBRARY"
]

# === SUBJECT NAME NORMALIZATION ===
SUBJECT_ALIASES = {
    "CN": "CN", "DEVOPS": "DEVOPS", "PPL": "PPL", "NLP": "NLP", "DAA": "DAA",
    "CN LAB": "CN LAB", "DEVOPS LAB": "DEVOPS LAB", "ACS LAB": "ACS LAB", "IPR": "IPR",
    "SPORTS": "SPORTS", "MEN": "MENTORING", "ASSOC": "ASSOCIATION", "ASSOCIATION": "ASSOCIATION",
    "LIB": "LIBRARY", "LIBRARY": "LIBRARY"
}

# === SETUP GOOGLE SHEETS CLIENT ROTATOR ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

def get_client_rotator():
    index = 0
    while True:
        cred_file = CREDENTIAL_FILES[index]
        if not os.path.exists(cred_file):
            raise FileNotFoundError(f"Credential file {cred_file} not found")
        creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, scope)
        yield gspread.authorize(creds)
        index = (index + 1) % len(CREDENTIAL_FILES)

client_rotator = get_client_rotator()
def next_client():
    return next(client_rotator)

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

# === ROLL NUMBERS ===
def generate_roll_numbers():
    rolls = [BASE_PREFIX + str(n) for n in range(72, 100) if str(n) not in ["80", "88"]]
    rolls += [BASE_PREFIX + f"{l}{d}" for l in "ABCD" for d in range(10)]
    return rolls

# === CLEAR ATTENDANCE SHEET ===
def clear_attendance_sheet(class_sheet):
    ranges = [
        "D8:D20", "J8:J20", "F27:F91", "H27:H91", "J27:J91", "L27:L91",
        "N27:N91", "P27:P91", "R27:R91", "T27:T91", "V27:V91", "X27:X91",
        "Z27:Z91", "AB27:AB91", "AD27:AD91"
    ]
    class_sheet.batch_clear(ranges)
    print("🧹 Cleared Attendence CSE-B(2023-27) ranges.")

# === ADD COLUMN TO SUBJECT SHEET ===
def prepare_new_column(sheet):
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, timestamp)
    return 3

# === MAP ROLL TO ROW INDEX IN SHEET ===
def get_roll_row_mapping(sheet):
    all_rows = sheet.get_all_values()
    return {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}

# === FETCH CLASSES HELD FOR ROLL ===
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
        driver.quit()

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
            driver.quit()
    print(f"❌ Failed to scrape {rollP}")
    return (rollP[:-1], {})

# === MAIN SCRAPING FUNCTION ===
def run_parallel_scraping():
    client = next_client()
    sheets = {name: client.open_by_key(SHEET_ID).worksheet(name) for name in SUBJECT_SHEETS}
    class_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

    clear_attendance_sheet(class_sheet)
    rolls = generate_roll_numbers()
    roll_with_p = [r + "P" for r in rolls]
    roll_to_row = {s: get_roll_row_mapping(sheets[s]) for s in SUBJECT_SHEETS}
    col_index = {s: prepare_new_column(sheets[s]) for s in SUBJECT_SHEETS}

    class_sheet.update("D8:D20", [[v] for v in extract_classes_held(roll_with_p[0])])
    class_sheet.update("J8:J20", [[v] for v in extract_classes_held("237Z1A05A8P")])

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {executor.submit(process_roll, r): r for r in roll_with_p}
        for f in as_completed(futures):
            roll, data = f.result()
            if not data:
                continue
            client = next_client()
            sheets = {name: client.open_by_key(SHEET_ID).worksheet(name) for name in SUBJECT_SHEETS}
            for subject, val in data.items():
                if roll in roll_to_row.get(subject, {}):
                    row = roll_to_row[subject][roll]
                    col = col_index[subject]
                    val = val if subject == "Overall %" else val + " %"
                    try:
                        sheets[subject].update_cell(row, col, val)
                        print(f"✅ Updated {subject} → {roll}: {val}")
                    except Exception as e:
                        print(f"❌ Error writing to {subject} for {roll}: {e}")
                else:
                    print(f"⚠️ Roll {roll} not found in sheet: {subject}")

# === RUN ===
if __name__ == "__main__":
    run_parallel_scraping()
