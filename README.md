# Outlook Lead Scraper for DocuSpeak

This script connects to your Outlook email, scans for emails from business domains (excluding Gmail, Yahoo, etc.), and exports the contact details to a CSV file and optionally to a Google Sheet.

## Prerequisites

1.  **Python 3.8+** installed.
2.  **Outlook Account** with IMAP enabled.
3.  **Google Cloud Project** (optional, for Sheets export).

## Setup

1.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

2.  **Configure Environment Variables:**
    Create a `.env` file in the root directory with the following content:

    ```env
    OUTLOOK_EMAIL=your_email@outlook.com
    OUTLOOK_PASSWORD=your_app_password
    SPREADSHEET_ID=your_google_sheet_id
    GOOGLE_CREDS_FILE=credentials.json
    ```

    *   `OUTLOOK_PASSWORD`: Do not use your regular password. Enable 2FA on your Microsoft account and generate an **App Password**.
    *   `SPREADSHEET_ID`: The ID found in the URL of your Google Sheet (e.g., `https://docs.google.com/spreadsheets/d/THIS_PART_IS_THE_ID/edit`).
    *   `GOOGLE_CREDS_FILE`: Path to your Google Service Account JSON key file.

3.  **Google Sheets Setup (Optional):**
    *   Go to the [Google Cloud Console](https://console.cloud.google.com/).
    *   Create a new project.
    *   Enable the **Google Sheets API** and **Google Drive API**.
    *   Create a **Service Account** and download the JSON key file. Rename it to `credentials.json` and place it in the project folder.
    *   Open your Google Sheet and share it with the `client_email` found inside your `credentials.json` file (give 'Editor' access).

## Usage

Run the script:

```bash
python scraper.py
```

The script will:
1.  Connect to Outlook.
2.  Fetch the last 500 emails (configurable in code).
3.  Filter out generic email domains.
4.  Save found leads to `leads.csv`.
5.  Upload leads to the specified Google Sheet.
