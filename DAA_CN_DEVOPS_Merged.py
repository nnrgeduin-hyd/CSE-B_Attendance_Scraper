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

# === CONFIGURATION ===
SUBJECT_CONFIGS = [
    {"name": "DAA", "cred": "credentials2.json", "main_col": 6},     # F
    {"name": "CN", "cred": "credentials3.json", "main_col": 8},      # H
    {"name": "DEVOPS", "cred": "credentials4.json", "main_col": 10}  # J
]

SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"
BASE_PREFIX = "237Z1A05"
THREADS = 15
MAX_ATTEMPTS = 3
BATCH_SIZE = THREADS

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

# === UTILITIES ===
def generate_roll_numbers():
    rolls = [BASE_PREFIX + str(num) for num in range(72, 100) if str(num) not in ["80", "88"]]
    rolls += [BASE_PREFIX + f"{letter}{d}" for letter in "ABCD" for d in range(10)]
    return rolls

def setup_gspread(cred_file):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, scope)
    client = gspread.authorize(creds)
    return client

def get_roll_row_mapping(sheet, col_range, start_row):
    rows = sheet.get(col_range)
    return {row[0].strip(): idx + start_row for idx, row in enumerate(rows) if row and row[0].strip()}

def get_roll_row_mapping_subject(sheet):
    all_rows = sheet.get_all_values()
    return {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}

def prepare_subject_column(sheet):
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, timestamp)
    return 3

def scrape_attendance(subject, rollP):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(40)
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
                if subject.upper() in subject_name:
                    percentage = cols[5].text.strip()
                    attended = cols[4].text.strip()
                    return (rollP[:-1], percentage, attended)
            return (rollP[:-1], None, None)
        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed for {rollP} ({subject}): {e}")
        finally:
            try:
                driver.quit()
            except:
                pass
    print(f"❌ Failed to scrape {rollP} ({subject})")
    return (rollP[:-1], None, None)

# === PER-SUBJECT SCRAPER ===
def process_subject(subject_cfg):
    subject = subject_cfg["name"]
    cred_file = subject_cfg["cred"]
    main_col = subject_cfg["main_col"]

    print(f"\n🚀 Starting subject: {subject}")
    client = setup_gspread(cred_file)
    subject_sheet = client.open_by_key(SHEET_ID).worksheet(subject)
    main_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

    main_map = get_roll_row_mapping(main_sheet, "B27:B91", 27)
    subj_map = get_roll_row_mapping_subject(subject_sheet)
    col_index = prepare_subject_column(subject_sheet)
    all_rolls = generate_roll_numbers()

    # Clear main sheet column for this subject
    main_col_letter = chr(64 + main_col)  # Convert 6 -> F, 8 -> H, 10 -> J
    main_sheet.batch_clear([f"{main_col_letter}27:{main_col_letter}91"])

    for i in range(0, len(all_rolls), BATCH_SIZE):
        batch = all_rolls[i:i + BATCH_SIZE]
        results = []

        with ThreadPoolExecutor(max_workers=THREADS) as executor:
            futures = [executor.submit(scrape_attendance, subject, roll + "P") for roll in batch]
            for future in as_completed(futures):
                result = future.result()
                if result:
                    results.append(result)

        subject_cells = []
        main_cells = []

        for roll, percent, attended in results:
            if percent and roll in subj_map:
                subject_cells.append(gspread.Cell(row=subj_map[roll], col=col_index, value=percent + " %"))
                print(f"✅ {subject} - {roll} => {percent}%")
            if attended and roll in main_map:
                main_cells.append(gspread.Cell(row=main_map[roll], col=main_col, value=attended))
            else:
                print(f"❌ {subject} - Missing roll {roll} in main sheet")

        if subject_cells:
            subject_sheet.update_cells(subject_cells)
            print(f"🟢 {subject} Sheet: {len(subject_cells)} % cells inserted")
        if main_cells:
            main_sheet.update_cells(main_cells)
            print(f"🟢 Main Sheet: {len(main_cells)} attended values inserted")

# === TOP LEVEL PARALLEL EXECUTION ===
if __name__ == "__main__":
    with ThreadPoolExecutor(max_workers=len(SUBJECT_CONFIGS)) as executor:
        futures = [executor.submit(process_subject, cfg) for cfg in SUBJECT_CONFIGS]
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as e:
                print(f"❌ Error in one subject thread: {e}")
