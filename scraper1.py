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
BASE_PREFIX = "237Z1A05"

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
overall_sheet = client.open_by_key(SHEET_ID).worksheet("Overall %")
daa_sheet = client.open_by_key(SHEET_ID).worksheet("DAA")

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
    print(f"✅ Found Chromium at: {chrome_path}")
else:
    print("⚠️ Chromium not found, using default Chrome")

# === Generate Roll Numbers ===
def generate_roll_numbers():
    rolls = []
    for num in range(72, 100):
        if str(num) in ["80", "88"]:
            continue
        rolls.append(BASE_PREFIX + str(num))
    for letter in ["A", "B", "C", "D"]:
        for d in range(0, 10):
            if f"{letter}{d}" == "A0":
                continue
            rolls.append(BASE_PREFIX + f"{letter}{d}")
    print(f"📋 Generated {len(rolls)} roll numbers")
    return rolls

# === Scraper ===
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

            # DAA Attendance from subject table
            daa_attendance = ""
            table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
            rows = table.find_elements(By.TAG_NAME, "tr")
            for row in rows[1:]:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 6:
                    subject_name = cols[1].text.strip()
                    if subject_name.startswith("DAA"):
                        daa_attendance = cols[5].text.strip()
                        break

            print(f"✅ {rollP} → Overall: {overall}, DAA: {daa_attendance}")
            return (rollP[:-1], overall, daa_attendance)

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed: {rollP} — {e}")
            time.sleep(0.5)
        finally:
            try:
                driver.quit()
            except:
                pass

    print(f"❌ Max attempts failed: {rollP}")
    return (rollP[:-1], "", "")

# === Insert new column with timestamp ===
def prepare_new_column(sheet_obj):
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    sheet_obj.insert_cols([[]], 3)
    sheet_obj.update_cell(10, 3, timestamp)
    print(f"📅 Created column C in '{sheet_obj.title}' with timestamp: {timestamp}")
    return 3

# === Roll mapping ===
def get_roll_row_mapping(sheet_obj):
    rows = sheet_obj.get_all_values()
    roll_map = {}
    for idx, row in enumerate(rows[10:], start=11):  # after header
        if row and row[0].strip():
            roll_map[row[0].strip()] = idx
    return roll_map

# === Run Scraping ===
def run_parallel_scraping():
    rolls = generate_roll_numbers()
    roll_map_overall = get_roll_row_mapping(overall_sheet)
    roll_map_daa = get_roll_row_mapping(daa_sheet)

    col_overall = prepare_new_column(overall_sheet)
    col_daa = prepare_new_column(daa_sheet)

    rolls_with_P = [r + "P" for r in rolls]
    total_batches = (len(rolls_with_P) + MAX_THREADS - 1) // MAX_THREADS

    for i in range(total_batches):
        batch = rolls_with_P[i * MAX_THREADS : (i + 1) * MAX_THREADS]
        print(f"\n🚀 Batch {i + 1}/{total_batches}: {batch}")

        results = {}
        with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
            futures = [executor.submit(process_roll, roll) for roll in batch]
            for future in as_completed(futures):
                roll, overall, daa = future.result()
                results[roll] = (overall, daa)

        for roll, (overall, daa) in results.items():
            if roll in roll_map_overall:
                overall_sheet.update_cell(roll_map_overall[roll], col_overall, overall)
            else:
                print(f"⚠️ {roll} not found in 'Overall %' sheet")

            if roll in roll_map_daa:
                daa_sheet.update_cell(roll_map_daa[roll], col_daa, daa)
            else:
                print(f"⚠️ {roll} not found in 'DAA' sheet")

        time.sleep(1)

    print("\n✅ All rolls processed & sheets updated!")

# === MAIN ===
if __name__ == "__main__":
    run_parallel_scraping()
