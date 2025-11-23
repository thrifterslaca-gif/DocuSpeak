import os
import csv
import re
from datetime import date
from imap_tools import MailBox, AND
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration
EMAIL_USER = os.getenv('OUTLOOK_EMAIL')
EMAIL_PASS = os.getenv('OUTLOOK_PASSWORD') # App Password recommended
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
GOOGLE_CREDS_FILE = os.getenv('GOOGLE_CREDS_FILE', 'credentials.json')
CSV_FILENAME = 'leads.csv'
IMAP_SERVER = 'outlook.office365.com'

# List of common public email domains to exclude
PUBLIC_DOMAINS = {
    'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com',
    'live.com', 'msn.com', 'icloud.com', 'aol.com', 'protonmail.com',
    'me.com', 'mac.com'
}

def is_business_email(email_address):
    """Check if the email domain is likely a business domain."""
    if not email_address or '@' not in email_address:
        return False
    domain = email_address.split('@')[1].lower()
    return domain not in PUBLIC_DOMAINS

def get_company_name(email_address):
    """Infer company name from email domain."""
    domain = email_address.split('@')[1]
    # Remove TLD (e.g., .com, .co.uk)
    parts = domain.split('.')
    if len(parts) > 2:
        # heuristic for things like .co.uk or subdomains
        # simply taking the first part as a guess, but often the second to last is better if it's not a generic co.uk
        # For simplicity, taking the part before the first dot, or if it's very short, maybe looking deeper.
        # But commonly: mail.google.com -> mail? no.
        # company.co.uk -> company
        company = parts[0]
    else:
        company = parts[0]
    return company.capitalize()

def connect_and_scrape():
    print("Connecting to Outlook...")

    if not EMAIL_USER or not EMAIL_PASS:
        print("Error: OUTLOOK_EMAIL and OUTLOOK_PASSWORD environment variables must be set.")
        return []

    leads = []

    try:
        with MailBox(IMAP_SERVER).login(EMAIL_USER, EMAIL_PASS) as mailbox:
            # Fetch emails. You can adjust criteria (e.g., specific folder, date)
            # Fetching all emails might be slow. Limiting to last 500 for demonstration.
            print("Fetching emails...")
            for msg in mailbox.fetch(limit=500, reverse=True):
                email_address = msg.from_
                name = msg.from_values.name

                # Clean up name if empty
                if not name:
                    name = email_address.split('@')[0]

                if is_business_email(email_address):
                    company = get_company_name(email_address)

                    lead = {
                        'Company': company,
                        'Name': name,
                        'Email': email_address,
                        'Date Found': date.today().isoformat()
                    }

                    # Avoid duplicates in the current run list
                    if not any(l['Email'] == email_address for l in leads):
                        leads.append(lead)
                        print(f"Found lead: {name} at {company}")
    except Exception as e:
        print(f"Error connecting to email: {e}")
        return []

    print(f"Total business leads found: {len(leads)}")
    return leads

def save_to_csv(leads):
    if not leads:
        print("No leads to save to CSV.")
        return

    # To avoid duplicates in CSV over time, one would ideally read the existing CSV.
    # But CSV is often just a dump. We will just overwrite or append?
    # Let's overwrite 'leads.csv' with the current run's finding, or maybe append?
    # User might prefer a fresh list or an accumulated one.
    # Let's assume this script is run to get a fresh batch or we can try to append if file exists.

    mode = 'w'
    header = True
    if os.path.exists(CSV_FILENAME):
        mode = 'a'
        header = False

    # However, if we append, we might duplicate.
    # For a simple script, let's just write the current batch to a new file or overwrite.
    # The user can manage the CSV.
    df = pd.DataFrame(leads)
    df.to_csv(CSV_FILENAME, index=False, mode=mode, header=header)
    print(f"Leads saved to {CSV_FILENAME}")

def save_to_google_sheets(leads):
    if not leads:
        print("No leads to save to Google Sheets.")
        return

    if not os.path.exists(GOOGLE_CREDS_FILE):
        print(f"Google credentials file '{GOOGLE_CREDS_FILE}' not found. Skipping Sheets export.")
        return

    if not SPREADSHEET_ID:
        print("SPREADSHEET_ID not set. Skipping Sheets export.")
        return

    print("Uploading to Google Sheets...")
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1

        # Check existing data to avoid duplicates
        existing_data = sheet.get_all_values()
        existing_emails = set()

        if existing_data:
            # Assuming 'Email' is the 3rd column (index 2) based on the order: Company, Name, Email, Date Found
            # Find the header index for 'Email' to be safe
            headers = existing_data[0]
            try:
                email_idx = headers.index('Email')
                for row in existing_data[1:]:
                    if len(row) > email_idx:
                        existing_emails.add(row[email_idx])
            except ValueError:
                # Header not found, maybe empty or different structure
                pass
        else:
             # If empty, add headers
            headers = list(leads[0].keys())
            sheet.append_row(headers)

        new_leads = [lead for lead in leads if lead['Email'] not in existing_emails]

        if new_leads:
            values = [list(lead.values()) for lead in new_leads]
            sheet.append_rows(values)
            print(f"Successfully uploaded {len(new_leads)} new leads to Google Sheets.")
        else:
            print("No new leads to upload (all duplicates).")

    except Exception as e:
        print(f"Error uploading to Google Sheets: {e}")

def main():
    leads = connect_and_scrape()
    if leads:
        save_to_csv(leads)
        save_to_google_sheets(leads)
    else:
        print("No leads found.")

if __name__ == "__main__":
    main()
