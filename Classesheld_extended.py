from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from zoneinfo import ZoneInfo
import gspread
import time
from shutil import which

# === CONFIG ===
SHEET_ID = "168dU0XLrRkVZQquAStktg_X9pMi3Vx9o9fOmbUYOUvA"

# === SETUP GOOGLE SHEETS ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
class_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

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

# === CLASSES HELD FOR ONE ROLL ===
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
        return held + ["0"] * (13 - len(held))  # pad to 13
    except Exception as e:
        print(f"❌ Error fetching classes held for {rollP}: {e}")
        return ["0"] * 13
    finally:
        try: driver.quit()
        except: pass

# === MAIN ===
def main():
    # → Classes Held for 237Z1A0572
    held_72 = extract_classes_held("237Z1A0572P")
    class_sheet.update("D8:D20", [[v] for v in held_72])
    print("✅ Inserted Classes Held from 72 into D8:D20")

    # → Classes Held for 237Z1A05A8
    held_a8 = extract_classes_held("237Z1A05A8P")
    class_sheet.update("J8:J20", [[v] for v in held_a8])
    print("✅ Inserted Classes Held from A8 into J8:J20")

if __name__ == "__main__":
    main()
