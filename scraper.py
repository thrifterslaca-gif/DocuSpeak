import os
import csv
import re
from datetime import date
from O365 import Account
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration
AZURE_CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
GOOGLE_CREDS_FILE = os.getenv('GOOGLE_CREDS_FILE', 'credentials.json')
CSV_FILENAME = 'leads.csv'

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
    # Simple heuristic: take the first part.
    # e.g. company.com -> company
    # company.co.uk -> company
    company = parts[0]
    return company.capitalize()

def connect_and_scrape():
    print("Connecting to Outlook via Microsoft Graph...")

    if not AZURE_CLIENT_ID:
        print("Error: AZURE_CLIENT_ID environment variable must be set.")
        print("Please see README.md for instructions on creating an Azure App.")
        return []

    leads = []

    try:
        # Authenticate with Device Code Flow or Interactive
        # For a local script without a secret, we act as a Public Client.
        credentials = (AZURE_CLIENT_ID, )
        account = Account(credentials)

        if not account.is_authenticated:
            print("Authentication required.")
            # 'basic' scope is usually enough for profile, 'message_all' for reading mail
            if account.authenticate(scopes=['basic', 'message_all']):
                print('Authenticated!')
            else:
                print("Authentication failed.")
                return []

        mailbox = account.mailbox()
        inbox = mailbox.get_folder(folder_name='Inbox')

        print("Fetching emails...")
        # Retrieve last 500 messages
        # O365 library returns a generator or iterable
        query = inbox.new_query().order_by('receivedDateTime', ascending=False)

        count = 0
        for msg in inbox.get_messages(limit=500, query=query, download_attachments=False):
            sender = msg.sender
            # sender is an object with name and address
            name = sender.name
            email_address = sender.address

            if not email_address:
                continue

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

            count += 1
            if count % 50 == 0:
                print(f"Processed {count} emails...")

    except Exception as e:
        print(f"Error connecting to email: {e}")
        return []

    print(f"Total business leads found: {len(leads)}")
    return leads

def save_to_csv(leads):
    if not leads:
        print("No leads to save to CSV.")
        return

    mode = 'w'
    header = True
    if os.path.exists(CSV_FILENAME):
        mode = 'a'
        header = False

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
            # Assuming 'Email' is the 3rd column (index 2)
            # Try to find header
            headers = existing_data[0]
            try:
                email_idx = headers.index('Email')
                for row in existing_data[1:]:
                    if len(row) > email_idx:
                        existing_emails.add(row[email_idx])
            except ValueError:
                pass
        else:
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
