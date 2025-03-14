import streamlit as st
import pandas as pd
import os
import time
from email_automator import get_gmail_service, create_message, send_email, get_user_email
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration Settings
SCOPES = [
    'openid',  # Added openid scope
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/userinfo.email',  # Changed scope for reading profile
    'https://www.googleapis.com/auth/userinfo.profile'  # Added scope for user info
]
DELAY_BETWEEN_EMAILS = 2  # seconds
DEFAULT_TEMPLATE = """Dear {name},

{message}

Best regards,
{sender_name}"""

DEFAULT_SUBJECT = "Message from {sender_name}"

# Page configuration
st.set_page_config(
    page_title="Email Automation Tool",
    page_icon="✉️",
    layout="wide",
    initial_sidebar_state="expanded"
)

def initialize_session_state():
    if 'credentials' not in st.session_state:
        st.session_state.credentials = None
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'emails_sent' not in st.session_state:
        st.session_state.emails_sent = 0
    if 'template' not in st.session_state:
        st.session_state.template = DEFAULT_TEMPLATE
    if 'subject' not in st.session_state:
        st.session_state.subject = DEFAULT_SUBJECT
    if 'available_variables' not in st.session_state:
        st.session_state.available_variables = []
    if 'user_email' not in st.session_state:
        st.session_state.user_email = None

def authenticate():
    try:
        st.session_state.credentials = get_gmail_service()
        if st.session_state.credentials:
            st.session_state.authenticated = True
            # Get user's email address
            st.session_state.user_email = get_user_email(st.session_state.credentials)
            return True
    except Exception as e:
        st.error(f"Authentication failed: {str(e)}")
        return False

def format_template_preview(template, sample_data):
    try:
        preview = template.format(**sample_data)
        return preview.replace('\n', '<br>')
    except KeyError as e:
        return f"⚠️ Template error: Missing variable {str(e)}"
    except Exception as e:
        return f"⚠️ Template error: {str(e)}"

def main():
    initialize_session_state()
    
    st.title("✉️ Email Automation Tool")
    
    # Sidebar for authentication and configuration
    with st.sidebar:
        st.header("Authentication")
        if not st.session_state.authenticated:
            if st.button("Connect to Gmail"):
                authenticate()
        else:
            st.success("✓ Connected to Gmail")
            st.info(f"Sending as: {st.session_state.user_email}")
        
        st.markdown("---")
        st.header("Configuration")
        sender_name = st.text_input("Your Name (for signature)")
        
        # Add footer to sidebar
        st.sidebar.markdown("---")
        st.sidebar.markdown("""
        Made by Saipavan  
        © 2024 All rights reserved
        """)
    
    # Main content area
    if not st.session_state.authenticated:
        st.warning("Please connect to Gmail using the sidebar button to get started.")
        return

    # File upload section
    st.header("1. Upload Contact List")
    uploaded_file = st.file_uploader("Upload your CSV file", type=['csv'])
    
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            if 'email' not in df.columns:
                st.error("CSV file must contain an 'email' column!")
                return
            
            # Store available variables from CSV
            st.session_state.available_variables = list(df.columns)
            
            st.success(f"✓ Successfully loaded {len(df)} contacts")
            
            # Email template section
            st.header("2. Email Configuration")
            
            # Available variables display
            st.subheader("Available Variables")
            for col in df.columns:
                st.code(f"{{{col}}}", language=None)
            st.code("{sender_name}", language=None)
            
            # Subject line configuration
            st.subheader("Email Subject")
            subject_template = st.text_input(
                "Subject Line Template",
                value=st.session_state.subject,
                help="Use variables from above"
            )
            
            # Template configuration
            st.subheader("Email Template")
            col1, col2 = st.columns([6, 6])
            
            with col1:
                st.markdown("### Edit Template")
                email_template = st.text_area(
                    "",
                    value=st.session_state.template,
                    height=400,
                )
                
                if email_template != st.session_state.template:
                    st.session_state.template = email_template
                
            with col2:
                st.markdown("### Live Preview")
                # Create sample data for preview
                sample_data = {var: f"[Sample {var}]" for var in st.session_state.available_variables}
                sample_data['sender_name'] = sender_name or "[Your Name]"
                
                preview = format_template_preview(email_template, sample_data)
                st.markdown(preview, unsafe_allow_html=True)
            
            # Send emails section
            st.header("3. Send Emails")
            
            if st.button("Start Sending Emails"):
                if not sender_name:
                    st.error("Please enter your name in the sidebar!")
                    return
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for index, row in df.iterrows():
                    try:
                        # Prepare variables for template
                        template_vars = row.to_dict()
                        template_vars['sender_name'] = sender_name
                        
                        # Format message and subject
                        try:
                            message_text = email_template.format(**template_vars)
                            subject = subject_template.format(**template_vars)
                        except KeyError as e:
                            st.error(f"Template error: Missing variable {str(e)}")
                            break
                        
                        # Send email using authenticated user's email
                        email = create_message(
                            st.session_state.user_email,  # Use authenticated email
                            row['email'],
                            subject,
                            message_text
                        )
                        send_email(st.session_state.credentials, 'me', email)
                        
                        # Update progress
                        st.session_state.emails_sent += 1
                        progress = (index + 1) / len(df)
                        progress_bar.progress(progress)
                        status_text.text(f"Sending email to {row['email']}... ({index + 1}/{len(df)})")
                        
                        time.sleep(DELAY_BETWEEN_EMAILS)
                        
                    except Exception as e:
                        st.error(f"Error sending email to {row['email']}: {str(e)}")
                        continue
                
                st.success(f"✓ Successfully sent {st.session_state.emails_sent} emails!")
                
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

    # Display stats
    if st.session_state.emails_sent > 0:
        st.sidebar.markdown("---")
        st.sidebar.header("Statistics")
        st.sidebar.metric("Emails Sent", st.session_state.emails_sent)

if __name__ == "__main__":
    main() 
    