"""
NSE Option Chain - Download & Email
=====================================
Step 1: Launch https://www.nseindia.com/ and maximize
Step 2: Click Option Chain link
Step 3: Click Download link → Excel file saved to Downloads folder
Step 4: Go to Downloads folder
Step 5: Find the Excel file (option-chain-ED-NIFTY-<today's date>)
Step 6: Send that Excel file to your email

Dependencies:
    pip install selenium webdriver-manager

SMTP Config:
    Fill in EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER below.
    For Gmail, use an App Password (not your main password):
    https://myaccount.google.com/apppasswords

Usage:
    python nse_download_email.py
"""

import os
import time
import glob
import smtplib
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

try:
    import undetected_chromedriver as uc
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver import ChromeOptions

except Exception as e:
    print("IMPORT ERROR:")
    print(e)
    raise

try:
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    sys.exit("Run: pip install webdriver-manager")


# ─────────────────────────────────────────────────────────────────────────────
# ✏️  CONFIGURATION — Fill these in before running
# ─────────────────────────────────────────────────────────────────────────────
import os

EMAIL_SENDER   = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")

DOWNLOADS_FOLDER = os.getcwd()
INDEX            = "NIFTY"   # Change to BANKNIFTY / FINNIFTY if needed
WAIT_FOR_DOWNLOAD_SECONDS = 30   # Max seconds to wait for the file to appear
# ─────────────────────────────────────────────────────────────────────────────


def build_driver(download_dir: str) :
    """STEP 1 — Launch Chrome maximized, set download folder."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")  # Use new headless mode for Chrome 109+
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--lang=en-US")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # Force Chrome to save downloads to our target folder (no Save-As dialog)
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)

    service = ChromeService(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=options)
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"},
    )
    print("[STEP 1] ✓ Browser launched and maximized")
    return driver


def open_nse_homepage(driver) -> None:
    """STEP 1 (continued) — Open NSE India homepage."""
    print("[STEP 1] Opening NSE homepage…")
    driver.get("https://www.nseindia.com/")
    time.sleep(5)
    print("[STEP 1] ✓ NSE homepage loaded")


def click_option_chain(driver) -> None:
    """STEP 2 — Click the Option Chain nav link."""
    print("[STEP 2] Looking for Option Chain link…")
    wait = WebDriverWait(driver, 20)

    SELECTORS = [
        "//a[normalize-space()='Option Chain']",
        "//a[contains(text(),'Option Chain')]",
        "//a[contains(@href,'option-chain')]",
    ]

    for xpath in SELECTORS:
        try:
            link = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            driver.execute_script("arguments[0].click();", link)
            time.sleep(5)
            print(f"[STEP 2] ✓ Option Chain page loaded  |  URL: {driver.current_url}")
            return
        except Exception:
            continue

    print("[STEP 2] ERROR: Could not click Option Chain link.")
    driver.quit()
    sys.exit(1)


def click_download(driver) -> None:
    """STEP 3 — Click the Download (CSV/Excel) link on the Option Chain page."""
    print("[STEP 3] Looking for Download link…")
    wait = WebDriverWait(driver, 20)

    SELECTORS = [
        # Text-based
        "//a[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'DOWNLOAD')]",
        "//button[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'DOWNLOAD')]",
        "//span[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'DOWNLOAD')]",
        # Class/ID hints
        "//*[contains(@class,'download')]",
        "//*[contains(@id,'download')]",
        "//*[contains(@title,'Download')]",
        # Icon-based (NSE uses fa-download icon next to link)
        "//i[contains(@class,'fa-download')]/..",
        "//img[contains(@src,'download')]/parent::a",
    ]

    for xpath in SELECTORS:
        try:
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            print(f"[STEP 3] Found Download element → clicking…")
            driver.execute_script("arguments[0].click();", element)
            time.sleep(3)
            print("[STEP 3] ✓ Download clicked")
            return
        except Exception:
            continue

    print("[STEP 3] ERROR: Could not find the Download link.")
    driver.quit()
    sys.exit(1)


def wait_for_excel_file(download_dir: str, timeout: int = WAIT_FOR_DOWNLOAD_SECONDS) -> str | None:
    """
    STEP 4 & 5 — Monitor Downloads folder for the NSE option chain Excel file.
    File name pattern: option-chain-ED-NIFTY-<DD-Mon-YYYY>.csv  or  .xlsx
    Returns the full path of the file if found, else None.
    """
    today = datetime.now()
    # NSE names the file like: option-chain-ED-NIFTY-16-Jun-2026
    date_str = today.strftime("%d-%b-%Y")        # e.g. 16-Jun-2026
    pattern_csv  = os.path.join(download_dir, f"option-chain-ED-{INDEX}-{date_str}.csv")
    pattern_xlsx = os.path.join(download_dir, f"option-chain-ED-{INDEX}-{date_str}.xlsx")
    # Fallback: any option-chain file downloaded today
    pattern_any  = os.path.join(download_dir, f"option-chain-ED-{INDEX}-*.csv")
    pattern_any2 = os.path.join(download_dir, f"option-chain-ED-{INDEX}-*.xlsx")

    print(f"[STEP 4] Watching Downloads folder: {download_dir}")
    print(f"[STEP 5] Looking for: option-chain-ED-{INDEX}-{date_str}.(csv/xlsx)")

    deadline = time.time() + timeout
    while time.time() < deadline:
        for pat in [pattern_csv, pattern_xlsx, pattern_any, pattern_any2]:
            matches = glob.glob(pat)
            # Exclude partial Chrome downloads (.crdownload)
            matches = [m for m in matches if not m.endswith(".crdownload")]
            if matches:
                # Pick most recently modified
                found = max(matches, key=os.path.getmtime)
                print(f"[STEP 5] ✓ File found: {found}")
                return found
        time.sleep(1)
        print(f"[STEP 5] Waiting… ({int(deadline - time.time())}s left)", end="\r")

    print(f"\n[STEP 5] ERROR: File not found in Downloads after {timeout}s.")
    return None


def send_email(file_path: str) -> None:
    """STEP 6 — Send the Excel/CSV file as an email attachment via Gmail SMTP."""
    print(f"\n[STEP 6] Preparing to send email…")

    if EMAIL_SENDER == "your_email@gmail.com":
        print("[STEP 6] ⚠️  EMAIL NOT CONFIGURED — please set EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER at the top of this script.")
        print(f"[STEP 6] File is saved at: {file_path}")
        return False

    filename   = os.path.basename(file_path)
    today_str  = datetime.now().strftime("%d %b %Y")
    subject    = f"NSE {INDEX} Option Chain — {today_str}"
    body       = (
        f"Hi,\n\n"
        f"Please find attached the NSE {INDEX} Option Chain data for {today_str}.\n\n"
        f"File: {filename}\n\n"
        f"This email was sent automatically by the NSE Option Chain scraper.\n"
    )

    msg = MIMEMultipart()
    msg["From"]    = EMAIL_SENDER
    msg["To"]      = EMAIL_RECEIVER
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    # Attach the file
    with open(file_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    # Send via Gmail SMTP
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())
        print(f"[STEP 6] ✓ Email sent successfully to {EMAIL_RECEIVER}")
        return True
    except smtplib.SMTPAuthenticationError:
        print("[STEP 6] ERROR: Gmail authentication failed.")
        print("         Make sure you're using an App Password, not your main password.")
        print("         Generate one at: https://myaccount.google.com/apppasswords")
        return False
    except Exception as e:
        print(f"[STEP 6] ERROR sending email: {e}")
        return False


def delete_file(file_path: str) -> None:
    """STEP 7 — Delete the file from Downloads after successful email send."""
    print(f"\n[STEP 7] Deleting file from Downloads: {os.path.basename(file_path)}")
    try:
        os.remove(file_path)
        print(f"[STEP 7] ✓ File deleted successfully from Downloads folder")
    except FileNotFoundError:
        print(f"[STEP 7] WARNING: File not found (may have already been deleted): {file_path}")
    except PermissionError:
        print(f"[STEP 7] ERROR: Permission denied — could not delete {file_path}")
    except Exception as e:
        print(f"[STEP 7] ERROR deleting file: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    driver = None
    try:
        # Steps 1–3: Browser automation
        driver = build_driver(DOWNLOADS_FOLDER)
        open_nse_homepage(driver)    # Step 1
        click_option_chain(driver)   # Step 2
        click_download(driver)       # Step 3

        # Steps 4–5: Wait and find the downloaded file
        excel_file = wait_for_excel_file(DOWNLOADS_FOLDER)

        if excel_file:
            # Step 6: Email it
            email_sent = send_email(excel_file)

            # Step 7: Delete file only if email was sent successfully
            if email_sent:
                delete_file(excel_file)
            else:
                print("[INFO] Email was not sent successfully — file kept in Downloads folder.")
        else:
            print("[INFO] Could not locate the downloaded file. Check your Downloads folder manually.")

    finally:
        if driver:
            driver.quit()
            print("\n[INFO] Browser closed.")


if __name__ == "__main__":
    main()
