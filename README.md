# Email Automation Tool

This tool allows you to send personalized emails automatically using Gmail based on spreadsheet data.

## Setup Instructions

1. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Set up Google Cloud Project and Enable Gmail API:
   - Go to the [Google Cloud Console](https://console.cloud.google.com)
   - Create a new project
   - Enable the Gmail API
   - Create OAuth 2.0 credentials (Desktop Application)
   - Download the credentials and save as `credentials.json` in the project directory

3. Create a `.env` file with your Gmail address:
   ```
   EMAIL_ADDRESS=your.email@gmail.com
   ```

4. Prepare your spreadsheet:
   - Required columns: email, name (additional columns can be used as template variables)
   - Save as CSV format

## Usage

1. Place your CSV file in the project directory
2. Update the email template in `config.py` if needed
3. Run the script:
   ```
   python email_automator.py
   ```

## Features

- Personalized email sending using Gmail
- Template-based email content
- CSV data support
- Secure OAuth2 authentication
- Rate limiting to avoid Gmail limits
- Error handling and logging 