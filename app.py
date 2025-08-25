from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from dotenv import load_dotenv
from datetime import datetime, timedelta
from send_email import send_excel_email
from io import BytesIO
import pandas as pd
import time
import sys
import os
import math
import re
import traceback


# --- Load environment variables ---
if load_dotenv():
    print("Loaded .env file successfully")
else:
    sys.exit("Error: .env file not found or could not be loaded.")


# --- Environment variables ---
USERNAME = os.getenv("USER")
PASSWORD = os.getenv("PASS")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_EMAIL_PASS = os.getenv("SENDER_EMAIL_PASS")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")
MEET_URL = os.getenv("MEET_URL")
FAILED_MEET_URL = os.getenv("FAILED_MEET_URL")
PROFILE_MEET_URL = os.getenv("PROFILE_MEET_URL")
DAY_FREQUENCY = int(os.getenv("DAY_FREQUENCY", 1))


def run(playwright):
    browser = None
    context = None
    page = None

    try:
        # --- Launch browser ---
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        # --- Login ---
        page.goto(MEET_URL)
        page.wait_for_load_state("networkidle")
        print("Opened login page")

        page.fill('input[name="email"]', USERNAME)
        page.fill('input[name="password"]', PASSWORD)
        page.click('button[type="submit"]')
        page.wait_for_timeout(5000)

        page.wait_for_load_state("networkidle", timeout=10000)
        print("Logged in successfully")

        # --- Navigate to failed meetings ---
        page.goto(FAILED_MEET_URL)
        page.wait_for_load_state("networkidle")
        print("Navigated to failed meetings page")

        # --- Select date filter ---
        dates = select_date(page, DAY_FREQUENCY)
        page.wait_for_load_state("networkidle")

        # --- Extract meeting data ---
        data = go_through_all_pages(page)

        if data:
            print("Report sent via email")
        else:
            print("No failed meetings data to recorded")
        
        # --- Send report via email ---
        send_excel_email(SENDER_EMAIL, RECEIVER_EMAIL, SENDER_EMAIL_PASS, data, dates)
        

    except PlaywrightTimeoutError as e:
        print(f"Timeout error occurred: {e}")
        traceback.print_exc()
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        traceback.print_exc()
    finally:
        # --- Ensure logout before closing ---
        try:
            if page and PROFILE_MEET_URL:
                page.goto(PROFILE_MEET_URL)
                page.wait_for_load_state("networkidle", timeout=5000)
                page.locator("button:has-text('Logout')").click()
                print("Logged out successfully")
        except Exception as e:
            print(f"Could not log out properly: {e}")
        finally:
            if context:
                context.close()
            if browser:
                browser.close()
            print("Browser closed safely")


def select_date(page, days_before):
    """Selects a date in the datepicker based on offset days."""
    target_date = datetime.today() - timedelta(days=days_before)
    day = target_date.day
    weekday = target_date.strftime("%A")
    month = target_date.strftime("%B")
    year = target_date.year

    # Ordinal suffix
    suffix = "th" if 4 <= day <= 20 or 24 <= day <= 30 else ["st", "nd", "rd"][day % 10 - 1]
    aria_label = f"Choose {weekday}, {month} {day}{suffix}, {year}"
    month_year_text = f"{month} {year}"

    page.locator('input[placeholder="Select start date"]').click()

    while not page.locator("h2.react-datepicker__current-month", has_text=month_year_text).is_visible():
        page.click("button.react-datepicker__navigation--previous")

    page.locator(f"div[aria-label='{aria_label}']").click()

    return str(target_date.date()), str(datetime.today().date())


def print_pagination_info(page):
    """Extracts meeting info from a single table page."""
    rows = page.locator("table tbody tr")
    meetings = []

    for i in range(rows.count()):
        cells = rows.nth(i).locator("td")
        meetings.append({
            "Meeting ID": cells.nth(0).inner_text().strip(),
            "Email": cells.nth(1).inner_text().strip(),
            "Meeting Type": cells.nth(2).inner_text().strip(),
            "Subject": cells.nth(3).inner_text().strip(),
            "Date Time": cells.nth(4).inner_text().strip()
        })
        time.sleep(0.1)

    return meetings


def go_through_all_pages(page):
    """Iterates over all pagination pages and collects meeting data."""
    if page.locator("text=No failed meetings found").is_visible():
        print("No failed meetings found")
        return None

    summary_text = page.locator("nav[aria-label='Table navigation'] span").nth(0).inner_text()
    match = re.search(r"(\d+)-(\d+)\s+of\s+(\d+)", summary_text)

    if not match:
        raise ValueError("Could not parse pagination summary")

    start, end, total = map(int, match.groups())
    per_page = end - start + 1
    total_pages = math.ceil(total / per_page)

    print(f"Total records: {total}, Per page: {per_page}, Pages: {total_pages}")

    all_data = []

    for page_num in range(1, total_pages + 1):
        page.wait_for_load_state("networkidle")
        print(f"Processing page {page_num} of {total_pages}")
        button = page.locator(f"nav[aria-label='Table navigation'] >> text='{page_num}'")
        button.click()
        page.wait_for_load_state("networkidle")
        all_data.extend(print_pagination_info(page))

    return convert_to_excel(all_data)


def convert_to_excel(data):
    """Converts list of dictionaries to Excel (in-memory buffer)."""
    df = pd.DataFrame(data)
    excel_buffer = BytesIO()
    df.to_excel(excel_buffer, index=False, engine="openpyxl")
    print(f"Data saved with {len(df)} rows")
    return excel_buffer


if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
