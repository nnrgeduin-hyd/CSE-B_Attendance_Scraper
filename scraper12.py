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

# === ALIASES ===
SUBJECT_ALIASES = {
    "CN": "CN", "DEVOPS": "DEVOPS", "PPL": "PPL", "NLP": "NLP", "DAA": "DAA",
    "CN LAB": "CN LAB", "DEVOPS LAB": "DEVOPS LAB", "ACS LAB": "ACS LAB", "IPR": "IPR",
    "SPORTS": "Sports", "MEN": "Men", "ASSOC": "Association", "ASSOCIATION": "Association",
    "LIB": "Library", "LIBRARY": "Library"
}

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheets = {name: client.open_by_key(SHEET_ID).worksheet(name) for name in SUBJECT_SHEETS}
class_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

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

# === Roll Number Generation ===
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

# === Column Prep and Clear CSE-B Sheet ===
def prepare_new_column(sheet):
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, timestamp)
    return 3

def clear_class_sheet_ranges():
    clear_ranges = [
        "D8:D20", "J8:J20",
        "F27:F91", "H27:H91", "J27:J91",
        "L27:L91", "N27:N91", "P27:P91", "R27:R91",
        "T27:T91", "V27:V91", "X27:X91", "Z27:Z91",
        "AB27:AB91", "AD27:AD91"
    ]
    for rng in clear_ranges:
        try:
            class_sheet.batch_clear([rng])
            time.sleep(1.5)  # Throttle clears
        except Exception as e:
            print(f"⚠️ Failed to clear {rng}: {e}")

# === Roll Mapping ===
def get_roll_row_mapping(sheet):
    all_rows = sheet.get_all_values()
    return {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}

# === Attendance Scraper ===
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
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) < 6:
                    continue
                subject_text = cols[1].text.upper().split(":")[0].strip()
                attendance_percent = cols[5].text.strip()

                key = SUBJECT_ALIASES.get(subject_text)
                if key and attendance_percent and attendance_percent != "&nbsp;":
                    subject_data[key] = attendance_percent

            return (rollP[:-1], subject_data)

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed for {rollP} — {e}")
            time.sleep(0.5)
        finally:
            try:
                driver.quit()
            except:
                pass
    return (rollP[:-1], {})

# === Extract Classes Held ===
def extract_classes_held(rollP):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 5)

            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(rollP)
            driver.find_element(By.ID, "btnNext").click()
            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(rollP)
            driver.find_element(By.ID, "btnSubmit").click()
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))
            rows = driver.find_elements(By.XPATH, "//table[@id='ctl00_cpStud_grdSubject']//tr")[1:]

            held_list = []
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 4:
                    held_list.append(cols[3].text.strip())

            return held_list
        except Exception as e:
            print(f"⚠️ Held extract failed for {rollP} — {e}")
            time.sleep(0.5)
        finally:
            try:
                driver.quit()
            except:
                pass
    return []

# === Main Runner ===
def run_parallel_scraping():
    rolls = generate_roll_numbers()
    roll_with_p = [r + "P" for r in rolls]
    clear_class_sheet_ranges()

    roll_to_row = {sheet: get_roll_row_mapping(sheets[sheet]) for sheet in SUBJECT_SHEETS}
    col_index = {sheet: prepare_new_column(sheets[sheet]) for sheet in SUBJECT_SHEETS}

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {executor.submit(process_roll, roll): roll for roll in roll_with_p}
        for future in as_completed(futures):
            roll, subj_data = future.result()
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
                        time.sleep(3.1)  # throttle writes
                    except Exception as e:
                        print(f"❌ Error writing to {subject} for {roll}: {e}")
                else:
                    print(f"⚠️ Roll {roll} not found in sheet: {subject}")

    # Write "Classes Held" for 72
    held_72 = extract_classes_held(roll_with_p[0])
    time.sleep(2.5)
    class_sheet.update("D8:D20", [[v] for v in held_72])
    print("✅ Inserted Classes Held from 72 into D8:D20")
    time.sleep(2.5)

    # Write "Classes Held" for A8
    held_a8 = extract_classes_held("237Z1A05A8P")
    time.sleep(2.5)
    class_sheet.update("J8:J20", [[v] for v in held_a8])
    print("✅ Inserted Classes Held from A8 into J8:J20")
    time.sleep(2.5)

# === Entry Point ===
if __name__ == "__main__":
    run_parallel_scraping()
