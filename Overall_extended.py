from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor, as_completed
from shutil import which
import gspread
import time
from datetime import datetime
from zoneinfo import ZoneInfo

# === CONFIG ===
SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"
MAX_ATTEMPTS = 3
MAX_THREADS = 15
BASE_PREFIX = "237Z1A05"  # ✅ Correct prefix

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials1.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID).worksheet("Overall %")

# === Setup Chrome Options ===
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--window-size=1280,800")

# ✅ Detect chromium-browser path for GitHub Actions
chrome_path = which("chromium-browser")
if chrome_path:
    print(f"✅ Found Chromium at: {chrome_path}")
    chrome_options.binary_location = chrome_path
else:
    print("⚠️ Chromium not found, will use default Chrome")

# === Generate Roll Numbers (72→99, A1→D9) ===
def generate_roll_numbers():
    rolls = []

    # Phase 1: numeric 72–99
    for num in range(72, 100):
        code = str(num)
        if code in ["80", "88"]:  # skip invalids
            continue
        rolls.append(BASE_PREFIX + code)

    # Phase 2: A1–D9
    for letter in ["A", "B", "C", "D"]:
        for d in range(0, 10):
            code = f"{letter}{d}"
            if code == "A0":  # skip invalid
                continue
            rolls.append(BASE_PREFIX + code)

    print(f"📋 Generated {len(rolls)} roll numbers")
    return rolls

# === Scraper Worker ===
def process_roll(rollP):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(50)
            wait = WebDriverWait(driver, 5)

            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")

            # Username = Password = Roll + P
            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(rollP)
            driver.find_element(By.ID, "btnNext").click()

            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(rollP)
            driver.find_element(By.ID, "btnSubmit").click()

            # Click Dashboard
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

            # Get Attendance
            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_lblTotalPercentage")))
            attendance = driver.find_element(By.ID, "ctl00_cpStud_lblTotalPercentage").text.strip()

            print(f"✅ {rollP} → {attendance}")
            return (rollP[:-1], attendance)  # remove P before storing

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed: {rollP} — {e}")
            time.sleep(0.5)
        finally:
            try:
                driver.quit()
            except:
                pass

    print(f"❌ Max attempts failed: {rollP}")
    return (rollP[:-1], "")

# === Prepare new column (after Name = Column C) ===
def prepare_new_column():
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    current_datetime = ist_time.strftime("%Y-%m-%d %I:%M %p")

    # Insert a column at position 3 (C)
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, current_datetime)  # header in row 10

    print(f"📅 Created new column C with timestamp: {current_datetime}")
    return 3  # always C for this run

# === Get existing roll → row mapping from sheet ===
def get_roll_row_mapping():
    all_rows = sheet.get_all_values()
    roll_map = {}
    for idx, row in enumerate(all_rows[10:], start=11):  # after header row
        if len(row) > 0 and row[0].strip():
            roll_map[row[0].strip()] = idx
    return roll_map

# === Run Scraping ===
def run_parallel_scraping():
    rolls = generate_roll_numbers()  # generate 72→99 + A1→D9
    roll_to_row = get_roll_row_mapping()
    print(f"🗂 Found {len(roll_to_row)} rolls mapped in sheet")

    # Create column header
    col_position = prepare_new_column()

    # Convert to login format (add P)
    rolls_with_P = [r + "P" for r in rolls]

    # Process in batches
    batch_size = MAX_THREADS
    total_batches = (len(rolls_with_P) + batch_size - 1) // batch_size

    for batch_index in range(total_batches):
        start = batch_index * batch_size
        end = start + batch_size
        batch_rolls = rolls_with_P[start:end]

        print(f"\n🚀 Batch {batch_index + 1}/{total_batches}: {batch_rolls}")

        batch_results = {}
        with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
            futures = {executor.submit(process_roll, roll): roll for roll in batch_rolls}
            for future in as_completed(futures):
                roll, attendance = future.result()
                batch_results[roll] = attendance

        # ✅ Update sheet for batch
        for roll, attendance in batch_results.items():
            if roll in roll_to_row:
                row_idx = roll_to_row[roll]
                sheet.update_cell(row_idx, col_position, attendance)
            else:
                print(f"⚠️ Roll {roll} not found in sheet → skipped")

        time.sleep(1)

    print("\n✅ All rolls processed & attendance updated!")

# === MAIN ===
if __name__ == "__main__":
    run_parallel_scraping()
