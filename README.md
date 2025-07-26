# Outlook Email Search for macOS

This project provides two different approaches to search Outlook emails from macOS:

## Option 1: Microsoft Graph API (Recommended)

### Setup Instructions:

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Register an Azure App:**
   - Go to [Azure Portal](https://portal.azure.com)
   - Create a new App Registration
   - Add redirect URI: `http://localhost:8080`
   - Grant `Mail.Read` permissions

3. **Get Access Token:**
   - Use the Microsoft Authentication Library (MSAL) to get tokens
   - Or use the Azure CLI: `az account get-access-token --resource https://graph.microsoft.com`

4. **Run the script:**
   ```bash
   python outlook_search_mac.py
   ```

## Option 2: Apple Mail + AppleScript

### Setup Instructions:

1. **Configure Apple Mail:**
   - Add your Outlook account to Apple Mail
   - Make sure Apple Mail is running

2. **Run the script:**
   ```bash
   python apple_mail_search.py
   ```

## Features

- Search emails by content or subject
- Display email details (subject, sender, date)
- Interactive search interface
- Error handling and user feedback

## Notes

- Option 1 requires Microsoft Graph API setup but works with any Outlook account
- Option 2 is simpler but requires Apple Mail to be configured with your Outlook account
- Both scripts are designed for macOS compatibility 