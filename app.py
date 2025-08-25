from playwright.sync_api import sync_playwright
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

if load_dotenv():
    print("Loaded .env file successfully")
else:
    sys.exit("Error: .env file not found or could not be loaded.")

USERNAME = os.getenv("USER")
PASS = os.getenv("PASS")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_EMAIL_PASS = os.getenv("SENDER_EMAIL_PASS")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")
MEET_URL = os.getenv("MEET_URL")
FAILED_MEET_URL = os.getenv("FAILED_MEET_URL")
PROFILE_MEET_URL = os.getenv("PROFILE_MEET_URL")
DAY_FREQUENCY = int(os.getenv("DAY_FREQUENCY", 1))

def run(playwright):
    browser = playwright.chromium.launch(headless=True, slow_mo=50,)
    context = browser.new_context()
    page = context.new_page()
    page.goto(MEET_URL)
    page.wait_for_load_state('networkidle')
    print("Opened Login page")
    
    # # Log in
    page.fill('input[name="email"]', USERNAME)
    page.fill('input[name="password"]', PASS)
    page.click('button[type="submit"]')
    time.sleep(5)
    page.wait_for_load_state('networkidle')

    # # Navigate to the failed meetings page
    page.goto(FAILED_MEET_URL)
    page.wait_for_load_state('networkidle')
    print("Navigated to failed meetings page")

    dates = select_date(page, DAY_FREQUENCY)

    page.wait_for_load_state('networkidle')

    data = go_through_all_pages(page)


    send_excel_email(SENDER_EMAIL, RECEIVER_EMAIL, SENDER_EMAIL_PASS, data, dates)


    #LOGOUT
    page.goto(PROFILE_MEET_URL)
    page.wait_for_load_state('networkidle')
    print("Navigated to profile page")
    page.locator("button:has-text('Logout')").click()
    print("Logged out successfully")

    context.close()
    browser.close()


def select_date(page, days_before):
    # Calculate target date
    target_date = datetime.today() - timedelta(days=days_before)

    # Get parts for aria-label
    day = target_date.day
    weekday = target_date.strftime("%A")   # e.g. Thursday
    month = target_date.strftime("%B")     # e.g. September
    year = target_date.year

    # --- Handle ordinal suffix (1st, 2nd, 3rd, etc.) ---
    if 4 <= day <= 20 or 24 <= day <= 30:
        suffix = "th"
    else:
        suffix = ["st", "nd", "rd"][day % 10 - 1]

    aria_label = f"Choose {weekday}, {month} {day}{suffix}, {year}"
    month_year_text = f"{month} {year}"  # e.g. "September 2025"

    page.locator('input[placeholder="Select start date"]').click()
    # Navigate until correct month/year visible
    while not page.locator("h2.react-datepicker__current-month", has_text=month_year_text).is_visible():
        page.click("button.react-datepicker__navigation--previous")

    # Click the correct date
    page.locator(f"div[aria-label='{aria_label}']").click()

    return str(target_date.date()), str(datetime.today().date())


def print_pagination_info(page):
    rows = page.locator("table tbody tr")
    meetings = []
    for i in range(rows.count()):
        cells = rows.nth(i).locator("td")
        meeting_id = cells.nth(0).inner_text().strip()
        email = cells.nth(1).inner_text().strip()
        meeting_type = cells.nth(2).inner_text().strip()
        subject = cells.nth(3).inner_text().strip()
        date_time = cells.nth(4).inner_text().strip()

        meetings.append({
            "Meeting ID": meeting_id,
            "Email": email,
            "Meeting Type": meeting_type,
            "Subject": subject,
            "Date Time": date_time
        })
        time.sleep(0.1)  # Small delay to ensure all data is captured

    return meetings


def go_through_all_pages(page):
    if page.locator("text=No failed meetings found").is_visible():
        print("No data found")
        return None
    # Step 1: Extract total records and per-page size
    summary_text = page.locator("nav[aria-label='Table navigation'] span").nth(0).inner_text()
    # Example: "Showing 21-30 of 43"
    match = re.search(r"(\d+)-(\d+)\s+of\s+(\d+)", summary_text)
    if not match:
        raise ValueError("Could not parse pagination summary")
    start, end, total = map(int, match.groups())
    per_page = end - start + 1
    total_pages = math.ceil(total / per_page)

    print(f"Total records: {total}, Per page: {per_page}, Pages: {total_pages}")

    fetched_data_final = []

    if total_pages == 1:
        fetched_data_final.append(print_pagination_info(page))
    else:
        # Step 2: Loop over each page number
        for page_num in range(1, total_pages + 1):
            page.wait_for_load_state("networkidle")
            print(f"Processing page {page_num}...")         

            # Click the page number button
            button = page.locator(f"nav[aria-label='Table navigation'] >> text='{page_num}'")
            button.click()
            page.wait_for_load_state("networkidle")
            page_data = print_pagination_info(page)
            fetched_data_final.append(page_data)

    return convert_to_excel(fetched_data_final)

def convert_to_excel(data):
    data = data
    flat_data = [record for sublist in data for record in sublist]
    # Convert to DataFrame
    df = pd.DataFrame(flat_data)

    excel_buffer = BytesIO()
    df.to_excel(excel_buffer, index=False, engine='openpyxl')
    print(f"Data saved with {len(df)} rows.")

    return excel_buffer

with sync_playwright() as playwright:
    run(playwright)