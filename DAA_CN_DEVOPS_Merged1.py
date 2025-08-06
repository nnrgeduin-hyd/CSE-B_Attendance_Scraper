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
SUBJECT_CONFIGS = {
    "DAA": {"cred": "credentials2.json", "main_col": 6},
    "CN": {"cred": "credentials3.json", "main_col": 8},
    "DEVOPS": {"cred": "credentials4.json", "main_col": 10}
}

SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"
BASE_PREFIX = "237Z1A05"
THREADS = 5
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

# === Setup ===
def setup_gspread(cred_file):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, scope)
    return gspread.authorize(creds)

def generate_roll_numbers():
    rolls = [BASE_PREFIX + str(n) for n in range(72, 100) if str(n) not in ["80", "88"]]
    rolls += [BASE_PREFIX + f"{l}{d}" for l in "ABCD" for d in range(10)]
    return rolls

# === Scraper ===
def scrape_all_subjects(rollP):
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

            result = {}
            rows = driver.find_element(By.ID, "ctl00_cpStud_grdSubject").find_elements(By.TAG_NAME, "tr")[1:]

            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) < 6:
                    continue
                subject_name = cols[1].text.upper().strip()
                percent = cols[5].text.strip()
                attended = cols[4].text.strip()
                for key in SUBJECT_CONFIGS:
                    if key in subject_name:
                        result[key] = {"percent": percent, "attended": attended}
            return (rollP[:-1], result)

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed for {rollP}: {e}")
        finally:
            try:
                driver.quit()
            except:
                pass
    return (rollP[:-1], {})

# === Main Logic ===
def main():
    roll_numbers = generate_roll_numbers()

    # Setup Sheets once per subject
    clients = {}
    subject_sheets = {}
    subject_maps = {}
    subject_columns = {}

    for subject, cfg in SUBJECT_CONFIGS.items():
        clients[subject] = setup_gspread(cfg["cred"])
        subject_sheet = clients[subject].open_by_key(SHEET_ID).worksheet(subject)
        subject_sheets[subject] = subject_sheet

        # Insert column for new attendance
        ist_time = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %I:%M %p")
        subject_sheet.insert_cols([[]], 3)
        subject_sheet.update_cell(10, 3, ist_time)
        subject_columns[subject] = 3

        all_rows = subject_sheet.get_all_values()
        subject_maps[subject] = {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}

    # Main sheet (common)
    main_client = setup_gspread("credentials2.json")  # any one is fine
    main_sheet = main_client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")
    main_map = {row[0].strip(): idx + 27 for idx, row in enumerate(main_sheet.get("B27:B91")) if row and row[0].strip()}

    # Clear all subject columns in main sheet
    for subject, cfg in SUBJECT_CONFIGS.items():
        col_letter = chr(64 + cfg["main_col"])
        main_sheet.batch_clear([f"{col_letter}27:{col_letter}91"])

    # Start scraping all rolls
    with ThreadPoolExecutor(max_workers=THREADS) as executor:
        futures = [executor.submit(scrape_all_subjects, roll + "P") for roll in roll_numbers]

        subject_updates = {subj: [] for subj in SUBJECT_CONFIGS}
        main_updates = {subj: [] for subj in SUBJECT_CONFIGS}

        for future in as_completed(futures):
            roll, data = future.result()
            if not data:
                print(f"❌ No data for {roll}")
                continue

            for subject, values in data.items():
                if subject in subject_maps and roll in subject_maps[subject]:
                    row = subject_maps[subject][roll]
                    subject_updates[subject].append(
                        gspread.Cell(row=row, col=subject_columns[subject], value=values["percent"] + " %")
                    )
                if roll in main_map:
                    main_updates[subject].append(
                        gspread.Cell(row=main_map[roll], col=SUBJECT_CONFIGS[subject]["main_col"], value=values["attended"])
                    )
                print(f"✅ {roll} - {subject}: {values['percent']}%")

    # Apply updates to all sheets
    for subject in SUBJECT_CONFIGS:
        if subject_updates[subject]:
            subject_sheets[subject].update_cells(subject_updates[subject])
            print(f"🟢 {subject} sheet updated: {len(subject_updates[subject])} cells")

        if main_updates[subject]:
            main_sheet.update_cells(main_updates[subject])
            print(f"🟢 Main sheet updated for {subject}: {len(main_updates[subject])} attended values")

if __name__ == "__main__":
    main()
