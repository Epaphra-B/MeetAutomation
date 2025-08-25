# Automated Failed Meetings Report Generator

This automates the extraction of failed meeting data from a web application using [Playwright](https://playwright.dev/python/) and sends a report via email in Excel format.  
It includes robust error handling, automatic login/logout, and configurable date filtering.

---

## Features

- Secure login using environment variables (`.env` file).
- Selects failed meetings within a configurable date range.
- Extracts tabular meeting data (meeting ID, email, type, subject, date).
- Sends the results as an Excel report to the configured recipient.
- Includes error handling and ensures safe logout.

---

### Python Libraries

Install dependencies with:

```bash
pip install -r requirements.txt
````

> You also need to install Playwright browsers:

```bash
playwright install
```

---

## Configuration

### 1. Environment Variables (`.env`)

Create a `.env` file in the project root:

```ini
# Authentication
USER=your_username
PASS=your_password

# Email settings
SENDER_EMAIL=sender@example.com
SENDER_EMAIL_PASS=sender_password
RECEIVER_EMAIL=receiver@example.com

# URLs
MEET_URL=https://example.com/login
FAILED_MEET_URL=https://example.com/failed-meetings
PROFILE_MEET_URL=https://example.com/profile

# Number of days before today to filter
DAY_FREQUENCY=1
```

### 2. Email Setup

The script uses a helper function `send_excel_email` (custom module you provide).
Ensure it is configured with proper SMTP details to send emails.

## Output

- If failed meetings exist: an Excel file is generated in-memory and sent via email.
- If no data is found: logs a message and sends email notifying no records.

---

## Error Handling

- All critical steps (login, navigation, extraction, email) are wrapped with error handling.
- Ensures browser closes safely even if errors occur.
- Provides detailed stack traces for debugging.

---

## Example Log Output

```plaintext
Loaded .env file successfully
Opened login page
Logged in successfully
Navigated to failed meetings page
Total records: 43, Per page: 10, Pages: 5
Processing page 1 of 5
Processing page 2 of 5
Data saved with 43 rows
Report sent via email
Logged out successfully
Browser closed safely
```

## Notes

- Ensure `.env` file is **never committed** to version control.
- If using Gmail, you may need to enable "App Passwords" or adjust SMTP settings.
- This script is meant for **internal automation** and should comply with your organization’s policies.

---
