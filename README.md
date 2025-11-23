# Outlook Lead Scraper for DocuSpeak

This script connects to your Outlook email using modern authentication (Microsoft Graph API), scans for emails from business domains (excluding Gmail, Yahoo, etc.), and exports the contact details to a CSV file and optionally to a Google Sheet.

## Prerequisites

1.  **Python 3.8+** installed.
2.  **Outlook Account** (Office 365 / Microsoft 365).
3.  **Google Cloud Project** (optional, for Sheets export).

## Setup

1.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

2.  **Register an Azure App (Required for Outlook Access):**
    Since App Passwords are disabled, you must use Modern Authentication.
    1.  Go to the [Azure Portal](https://portal.azure.com/) and sign in.
    2.  Search for **"App registrations"** and click it.
    3.  Click **"New registration"**.
        *   **Name**: "DocuSpeak Scraper" (or anything you like).
        *   **Supported account types**: "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)".
        *   **Redirect URI**: Select **"Public client/native (mobile & desktop)"** and enter `https://login.microsoftonline.com/common/oauth2/nativeclient`.
        *   Click **Register**.
    4.  Copy the **Application (client) ID** from the Overview page. You will need this for the `.env` file.
    5.  **Note:** You do **not** need a Client Secret for a public desktop app.

3.  **Configure Environment Variables:**
    Create a `.env` file in the root directory with the following content:

    ```env
    AZURE_CLIENT_ID=your_application_client_id_from_azure
    SPREADSHEET_ID=your_google_sheet_id
    GOOGLE_CREDS_FILE=credentials.json
    ```

    *   `AZURE_CLIENT_ID`: The ID you copied in step 2.
    *   `SPREADSHEET_ID`: The ID found in the URL of your Google Sheet.
    *   `GOOGLE_CREDS_FILE`: Path to your Google Service Account JSON key file (see below).

4.  **Google Sheets Setup (Optional):**
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

1.  The first time you run it, the script will print a URL to the console.
2.  Copy and paste this URL into your browser.
3.  Sign in with your Outlook account and grant the requested permissions.
4.  Copy the **result URL** (the page you are redirected to, which will be a blank page or error page starting with `https://login.microsoftonline.com/...`) back into the console prompt.
5.  The script will save a token locally (in `o365_token.txt`) so you don't have to login again next time.

The script will then:
1.  Fetch the last 500 emails.
2.  Filter out generic email domains.
3.  Save found leads to `leads.csv`.
4.  Upload leads to the specified Google Sheet.
