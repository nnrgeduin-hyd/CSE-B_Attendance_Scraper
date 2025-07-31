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

ATTENDED_RANGES = {
    "DAA": "F27:F91", "CN": "H27:H91", "DEVOPS": "J27:J91", "PPL": "L27:L91",
    "NLP": "N27:N91", "CN LAB": "P27:P91", "DEVOPS LAB": "R27:R91", "ACS LAB": "T27:T91", "IPR": "V27:V91",
    "SPORTS": "X27:X91", "MEN": "Z27:Z91", "ASSOC": "AB27:AB91", "LIB": "AD27:AD91"
}

# === MAPPING FOR FUZZY SUBJECT MATCHING ===
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

# === Generate Roll Numbers ===
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

# === Insert New Column with Timestamp ===
def prepare_new_column(sheet):
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, timestamp)
    return 3

# === Clear Specific Ranges in Main Sheet ===
def clear_main_sheet():
    ranges = ["D8:D20", "J8:J20", "D27:D91", "F27:F91", "H27:H91", "J27:J91", "L27:L91",
              "N27:N91", "P27:P91", "R27:R91", "T27:T91", "V27:V91", "X27:X91", "Z27:Z91",
              "AB27:AB91", "AD27:AD91"]
    for rng in ranges:
        main_sheet.batch_clear([rng])
        time.sleep(1)

# === Get Roll Row Mapping ===
def get_roll_row_mapping(sheet):
    title = sheet.title.strip().upper()
    if title == "ATTENDENCE CSE-B(2023-27)":
        rolls = sheet.get("B27:B91")
        return {row[0].strip(): idx for idx, row in enumerate(rolls, start=27) if row and row[0].strip()}
    else:
        all_rows = sheet.get_all_values()
        return {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}

# === Scrape Attendance for Roll ===
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
            attended_data = {}
            held_data = {}

            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) < 6:
                    continue
                subject_text = cols[1].text.upper().split(":" if ":" in cols[1].text else " ")[0].strip()
                key = SUBJECT_ALIASES.get(subject_text)

                if key:
                    percent = cols[5].text.strip()
                    attended = cols[3].text.strip()
                    held = cols[2].text.strip()
                    if percent and percent != "&nbsp;":
                        subject_data[key] = percent
                    if attended.isdigit():
                        attended_data[key] = attended
                    if held.isdigit():
                        held_data[key] = held

            return (rollP[:-1], subject_data, attended_data, held_data)

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed for {rollP} — {e}")
            time.sleep(1.5)
        finally:
            try:
                driver.quit()
            except:
                pass

    print(f"❌ Failed to scrape {rollP}")
    return (rollP[:-1], {}, {}, {})

# === Run Main Logic ===
def run_parallel_scraping():
    rolls = generate_roll_numbers()
    roll_with_p = [r + "P" for r in rolls]

    roll_to_row = {sheet: get_roll_row_mapping(sheets[sheet]) for sheet in SUBJECT_SHEETS}
    roll_to_row_main = get_roll_row_mapping(main_sheet)
    col_index = {sheet: prepare_new_column(sheets[sheet]) for sheet in SUBJECT_SHEETS}
    clear_main_sheet()

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {executor.submit(process_roll, roll): roll for roll in roll_with_p}
        for future in as_completed(futures):
            roll, subj_data, attended_data, held_data = future.result()
            if not subj_data:
                continue

            for subject, value in subj_data.items():
                if roll in roll_to_row.get(subject, {}):
                    row = roll_to_row[subject][roll]
                    col = col_index[subject]
                    val = value if subject == "Overall %" else value + " %"
                    try:
                        sheets[subject].update_cell(row, col, val)
                        time.sleep(3.1)
                    except Exception as e:
                        print(f"❌ Write error {subject} {roll}: {e}")

            for subject, attended in attended_data.items():
                key = subject.upper()
                range_ref = ATTENDED_RANGES.get(key)
                if range_ref and roll in roll_to_row_main:
                    row_idx = roll_to_row_main[roll] - 27
                    try:
                        cell_range = main_sheet.range(range_ref)
                        cell_range[row_idx].value = attended
                        main_sheet.update_cells(cell_range)
                        time.sleep(2.5)
                    except Exception as e:
                        print(f"❌ Attend write error {subject} {roll}: {e}")

            if roll == "237Z1A0572" or roll == "237Z1A05A8":
                target_range = "D8:D20" if roll == "237Z1A0572" else "J8:J20"
                try:
                    held_list = [held_data.get(sub, "") for sub in ["DAA", "CN", "DEVOPS", "PPL", "NLP",
                                                                    "CN LAB", "DEVOPS LAB", "ACS LAB", "IPR",
                                                                    "SPORTS", "MEN", "ASSOC", "LIB"]]
                    cell_range = main_sheet.range(target_range)
                    for idx, val in enumerate(held_list):
                        cell_range[idx].value = val
                    main_sheet.update_cells(cell_range)
                    time.sleep(2.5)
                except Exception as e:
                    print(f"❌ Held write error {roll}: {e}")

if __name__ == "__main__":
    run_parallel_scraping()
