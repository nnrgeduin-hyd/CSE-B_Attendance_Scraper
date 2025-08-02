from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor, as_completed
from shutil import which
from datetime import datetime
import gspread
from zoneinfo import ZoneInfo

# === CONFIG ===
SUBJECT = "CN"
SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"
CREDENTIAL_FILE = "credentials.json"
MAX_ATTEMPTS = 3
BASE_PREFIX = "237Z1A05"
THREADS = 15
BATCH_SIZE = THREADS  # one batch = max thread count

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIAL_FILE, scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID).worksheet(SUBJECT)
attended_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

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

# === Helpers ===
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

def get_roll_row_mapping():
    all_rows = sheet.get_all_values()
    return {row[0].strip(): idx for idx, row in enumerate(all_rows[10:], start=11) if row and row[0].strip()}

def get_attended_row_map():
    data = attended_sheet.get_all_values()
    return {row[1].strip(): idx for idx, row in enumerate(data[26:], start=27) if len(row) > 1 and row[1].strip()}

def prepare_column():
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    timestamp = ist_time.strftime("%Y-%m-%d %I:%M %p")
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, timestamp)
    return 3

def scrape_attendance(rollP):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(30)
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
                if SUBJECT in subject_name:
                    attended = cols[3].text.strip()
                    return (rollP[:-1], attended if attended != "&nbsp;" else None)
            return (rollP[:-1], None)

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed for {rollP}: {e}")
        finally:
            try:
                driver.quit()
            except:
                pass
    print(f"❌ Failed to scrape {rollP}")
    return (rollP[:-1], None)

# === Main Execution ===
def main():
    roll_map = get_roll_row_mapping()
    col_index = prepare_column()
    all_rolls = generate_roll_numbers()

    # Clear H27:H91 in Attendence CSE-B(2023-27)
    attended_sheet.batch_clear(["H27:H91"])
    print("🧹 Cleared H27:H91 in 'Attendence CSE-B(2023-27)'")

    for i in range(0, len(all_rolls), BATCH_SIZE):
        batch = all_rolls[i:i + BATCH_SIZE]
        results = []

        with ThreadPoolExecutor(max_workers=THREADS) as executor:
            futures = [executor.submit(scrape_attendance, roll + "P") for roll in batch]
            for future in as_completed(futures):
                result = future.result()
                if result:
                    results.append(result)

        # Insert % into subject sheet
        batch_cells = []
        for roll, attendance in results:
            if not attendance:
                print(f"⚠️ No data for {roll}")
                continue
            if roll in roll_map:
                row = roll_map[roll]
                cell = gspread.Cell(row=row, col=col_index, value=attendance + " %")
                batch_cells.append(cell)
                print(f"✅ {roll} => {attendance}%")
            else:
                print(f"❌ Roll not in subject sheet: {roll}")
        if batch_cells:
            sheet.update_cells(batch_cells)
            print(f"🟢 Inserted % for {len(batch_cells)} rolls.")

        # Insert Classes Attended into Attendence CSE-B(2023-27)
        attended_map = get_attended_row_map()
        attended_col_index = 8  # Column H
        batch_cells_attended = []

        for roll, attended in results:
            if roll in attended_map and attended:
                row = attended_map[roll]
                cell = gspread.Cell(row=row, col=attended_col_index, value=attended)
                batch_cells_attended.append(cell)
                print(f"📘 {roll} => Attended: {attended}")

        if batch_cells_attended:
            attended_sheet.update_cells(batch_cells_attended)
            print(f"🟦 Inserted attended values in H27:H91 for {len(batch_cells_attended)} rolls.")

if __name__ == "__main__":
    main()
