import imaplib
import email
import re
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# ----------------- CONFIG -----------------
IMAP_SERVER = "mail.b2bgrowthexpo.com"
IMAP_USER = "speakermanagement@b2bgrowthexpo.com"
IMAP_PASS = "}g3=ynks717Z"

SENDER_FILTER = "The-Business-Revival-Series-2023@showoff.asp.events"

GOOGLE_SHEET_NAME = "Expo-Sales-Management"
SHEET_TAB = "speakers-2"

# Google service account credentials
SERVICE_ACCOUNT_FILE = "/etc/secrets/credentials.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
CREDS = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# ----------------- CONNECT SHEET -----------------
client = gspread.authorize(CREDS)
sheet = client.open(GOOGLE_SHEET_NAME).worksheet(SHEET_TAB)


# ----------------- CLEANER -----------------
def clean_text(text):
    if not text:
        return ""
    return text.replace("&nbsp;", " ").replace("\xa0", " ").strip()


# ----------------- PARSE EMAIL BODY -----------------
def parse_details(body):
    details = {}
    patterns = {
        "First Name": r"First Name:\s*(.*)",
        "Last Name": r"Last Name:\s*(.*)",
        "Email": r"Email:\s*(.*)",
        "Mobile Number": r"Mobile Number:\s*(.*)",
        "LinkedIn Profile Link": r"LinkedIn Profile Link:\s*(.*)",
        "Business Name": r"Business Name:\s*(.*)",
        "Business linkedln page or Website": r"Business linkedln page or Website:\s*(.*)",
        "Which event are you interested in": r"Which event are you interested in:\s*(.*)",
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, body)
        details[key] = clean_text(match.group(1)) if match else ""

    return details


# ----------------- DUPLICATE CHECK -----------------
def get_existing_emails():
    """Fetch all emails from sheet for duplicate check."""
    try:
        all_emails = sheet.col_values(13)  # "Email" column index (13th)
        return set([e.lower().strip() for e in all_emails if e])  # normalize
    except Exception as e:
        print(f"❌ Error fetching existing emails: {e}")
        return set()


# ----------------- FETCH EMAILS -----------------
def fetch_emails():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(IMAP_USER, IMAP_PASS)
    mail.select("inbox")

    status, messages = mail.search(None, f'FROM "{SENDER_FILTER}"')
    email_ids = messages[0].split()

    leads = []
    for eid in email_ids[-50:]:  # check last 50 emails
        _, msg_data = mail.fetch(eid, "(RFC822)")
        raw_msg = msg_data[0][1]
        msg = email.message_from_bytes(raw_msg)

        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    leads.append(parse_details(body))
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")
            leads.append(parse_details(body))

    mail.logout()
    return leads


# ----------------- PROCESS EMAILS (BATCH) -----------------
def process_emails(leads):
    existing_emails = get_existing_emails()
    new_rows = []

    for details in leads:
        email_value = details["Email"].lower().strip()

        if not email_value:
            print("⚠️ No email found, skipping.")
            continue

        if email_value in existing_emails:
            print(f"⏩ Duplicate skipped: {email_value}")
            continue

        row = [
            datetime.now().strftime("%Y-%m-%d"),  # Lead Date
            "Businessrevivalseries",  # Lead Source
            details["First Name"],  # First_Name
            details["Last Name"],  # Last Name
            "",  # Email Sent-Date
            "",  # Reply Status
            details["Business Name"],  # Company Name
            "",  # Designation
            "Speaker_opportunity",  # Interested In?
            "",  # Comments
            "",  # Next Followup
            details["Mobile Number"],  # Mobile
            email_value,  # Email
            details["Which event are you interested in"],  # Show
            details["Business linkedln page or Website"],  # Company Linkedin Page
            details["LinkedIn Profile Link"],  # Personal Linkedin Page
        ]

        new_rows.append(row)
        existing_emails.add(email_value)  # <-- update immediately

    if new_rows:
        for row in reversed(new_rows):  # reverse so first email ends up closest to header
            sheet.insert_row(row, 2, value_input_option="USER_ENTERED")
        print(f"✅ Added {len(new_rows)} new leads",flush=True)
    else:
        print("ℹ️ No new leads to add",flush=True)


# ----------------- RUN -----------------
if __name__ == "__main__":
    leads = fetch_emails()
    process_emails(leads)
