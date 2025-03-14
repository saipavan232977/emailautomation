import os
import base64
import time
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
from dotenv import load_dotenv
import requests
from googleapiclient.discovery import build

# Load environment variables
load_dotenv()

# Gmail API configuration
SCOPES = [
    'openid',  # Added openid scope
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/userinfo.email',  # Changed scope for reading profile
    'https://www.googleapis.com/auth/userinfo.profile'  # Added scope for user info
]

# Set up logging
logging.basicConfig(
    filename="email_logs.txt",
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def get_gmail_service():
    """Set up and return Gmail API credentials."""
    creds = None
    
    # Load existing credentials if available
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # Refresh or create new credentials if needed
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for future use
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def create_message(sender, to, subject, message_text):
    """Create an email message."""
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    raw_message = base64.urlsafe_b64encode(message.as_bytes())
    return {'raw': raw_message.decode('utf-8')}

def send_email(creds, user_id, message):
    """Send an email message."""
    try:
        access_token = creds.token
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        response = requests.post(
            'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
            headers=headers,
            json=message
        )
        
        if response.status_code == 200:
            result = response.json()
            logging.info(f"Message Id: {result['id']} sent successfully")
            return result
        else:
            error_msg = f"Failed to send email. Status code: {response.status_code}, Response: {response.text}"
            logging.error(error_msg)
            raise Exception(error_msg)
            
    except Exception as error:
        logging.error(f'An error occurred: {error}')
        raise error

def process_spreadsheet(filename):
    """Read and process the spreadsheet data."""
    try:
        df = pd.read_csv(filename)
        required_columns = ['email', 'name']
        
        if not all(col in df.columns for col in required_columns):
            raise ValueError(f"Spreadsheet must contain columns: {required_columns}")
        
        return df
    except Exception as e:
        logging.error(f'Error processing spreadsheet: {e}')
        raise e

def get_user_email(credentials):
    """Get the email address of the authenticated user using OAuth2 userinfo endpoint."""
    try:
        access_token = credentials.token
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get('https://www.googleapis.com/oauth2/v2/userinfo', headers=headers)
        
        if response.status_code == 200:
            user_info = response.json()
            return user_info['email']
        else:
            error_msg = f"Failed to get user info. Status code: {response.status_code}, Response: {response.text}"
            raise Exception(error_msg)
            
    except Exception as e:
        raise Exception(f"Failed to get user email: {str(e)}")

def main():
    """Main function to run the email automation."""
    try:
        # Get email address from environment variables
        sender_email = os.getenv('EMAIL_ADDRESS')
        if not sender_email:
            raise ValueError("EMAIL_ADDRESS not found in .env file")

        # Set up Gmail API service
        creds = get_gmail_service()
        
        # Get CSV filename from user
        csv_file = input("Enter the name of your CSV file (e.g., contacts.csv): ")
        
        # Get sender name
        sender_name = input("Enter your name (for email signature): ")
        
        # Process spreadsheet
        df = process_spreadsheet(csv_file)
        
        # Initialize counters
        emails_sent = 0
        
        # Process each row in the spreadsheet
        for index, row in df.iterrows():
            if emails_sent >= MAX_EMAILS_PER_DAY:
                logging.warning("Daily email limit reached")
                break
                
            try:
                # Create personalized message
                custom_message = input(f"Enter custom message for {row['name']} (or press Enter to skip): ")
                if not custom_message:
                    continue
                
                message_text = EMAIL_TEMPLATE.format(
                    name=row['name'],
                    custom_message=custom_message,
                    sender_name=sender_name
                )
                
                subject = EMAIL_SUBJECT.format(name=row['name'])
                
                # Create and send email
                email = create_message(sender_email, row['email'], subject, message_text)
                send_email(creds, 'me', email)
                
                emails_sent += 1
                logging.info(f"Email sent to {row['email']}")
                
                # Delay between emails
                time.sleep(DELAY_BETWEEN_EMAILS)
                
            except Exception as e:
                logging.error(f"Error sending email to {row['email']}: {e}")
                continue
                
        print(f"Process completed. {emails_sent} emails sent successfully.")
        
    except Exception as e:
        logging.error(f"Application error: {e}")
        print(f"An error occurred. Please check the logs in {LOG_FILE}")

if __name__ == "__main__":
    main() 