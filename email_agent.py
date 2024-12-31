import time
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import base64
from docx import Document
import streamlit as st
import uuid
from datetime import datetime

# Constants
SCOPES = ['https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.readonly']
DEFAULT_TIME_INTERVALS = [30, 90, 120]  # Default time intervals in seconds

# Function to generate a unique log file name
def generate_log_filename():
    """Generate a unique log file name using timestamp and UUID."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    unique_id = uuid.uuid4().hex
    return f"Email_Logging_{timestamp}_{unique_id}.xlsx"

# Step 1: Load Email Templates from Word Document
def load_email_templates(docx_path):
    """Load email templates from a Word document with structured format."""
    try:
        doc = Document(docx_path)
        templates = []
        current_template = {}

        for para in doc.paragraphs:
            text = para.text.strip()  # Clean up the paragraph text
            if not text:  # Skip empty lines
                continue

            # Detect and handle Email_subject
            if text.lower().startswith("email_subject:"):
                if current_template:  # Save the previous template if it exists
                    templates.append(current_template)
                current_template = {"subject": text.split(":", 1)[1].strip(), "body": ""}

            # Detect and handle Email_body
            elif text.lower().startswith("email_body:"):
                current_template["body"] = ""

            # Add lines to the email body
            else:
                if "body" in current_template:
                    current_template["body"] += f"{text}\n"

        # Append the last template if exists
        if current_template:
            templates.append(current_template)

        return templates
    except Exception as e:
        st.error(f"Error occurred while reading Word document: {e}")
        raise

# Step 2: Gmail Authentication
def authenticate_gmail():
    """Authenticate and initialize Gmail API."""
    try:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=6231)
        return build('gmail', 'v1', credentials=creds)
    except Exception as e:
        st.error(f"Error occurred during Gmail authentication: {e}")
        raise

# Step 3: Send Email to a Single Recipient
def send_email(service, email_id, subject, body):
    """Send an email to a single recipient using Gmail API."""
    try:
        message = MIMEMultipart()
        message['to'] = email_id
        message['subject'] = subject
        message.attach(MIMEText(body))
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        sent_message = service.users().messages().send(userId="me", body={'raw': raw_message}).execute()
        return sent_message['id']  # Return the message ID to track the thread
    except HttpError as error:
        st.error(f"HttpError occurred while sending the email: {error}")
        return None
    except Exception as e:
        st.error(f"Unexpected error occurred while sending the email: {e}")
        raise

# Step 4: Check for Replies
def check_for_replies(service, sent_message_id, recipient_email):
    """Check for replies to the sent message."""
    if not sent_message_id:
        st.warning("No message ID found to track replies.")
        return False

    try:
        sent_message = service.users().messages().get(userId="me", id=sent_message_id).execute()
        thread_id = sent_message.get('threadId')

        if not thread_id:
            st.warning(f"No thread ID found for the sent message with ID {sent_message_id}")
            return False

        messages = service.users().threads().get(userId='me', id=thread_id).execute().get('messages', [])
        for message in messages:
            headers = {header['name']: header['value'] for header in message['payload']['headers']}
            sender = headers.get('From', '')
            if recipient_email in sender:
                st.success(f"Reply detected from {recipient_email}.")
                return True

        st.info("No reply detected in the thread.")
        return False
    except HttpError as error:
        st.error(f"HttpError occurred while checking for replies: {error}")
        return False
    except Exception as e:
        st.error(f"Unexpected error occurred while checking for replies: {e}")
        return False

# Step 5: Update Logs
def update_logs(log_filename, email_id, status):
    """Update the log file."""
    try:
        # Check if the log file exists, if not create it
        try:
            df = pd.read_excel(log_filename)
        except FileNotFoundError:
            df = pd.DataFrame(columns=["Email ID", "Status", "Timestamp"])

        new_entry = {"Email ID": email_id, "Status": status, "Timestamp": pd.Timestamp.now()}
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        df.to_excel(log_filename, index=False)
    except Exception as e:
        st.error(f"Error occurred while updating the log file: {e}")
        raise

# Step 6: Email Pipeline for Multiple Recipients and Different Company Names
def email_pipeline(service, recipients, email_templates, log_filename, time_intervals):
    """Pipeline to send periodic emails to different recipients with different company names and check for replies."""
    for recipient in recipients:
        email_id = recipient['email_id']
        company_name = recipient['company_name']

        st.info(f"Starting email pipeline for {email_id} - {company_name}...")

        sent_message_id = None
        for i, template in enumerate(email_templates):
            try:
                subject = template["subject"].format(company_name=company_name)
                body = template["body"].format(company_name=company_name)

                sent_message_id = send_email(service, email_id, subject, body)
                if not sent_message_id:
                    st.error(f"Error: Could not send email to {email_id}.")
                    update_logs(log_filename, email_id, "Email Sending Failed")
                    break

                st.info(f"Email {i+1} sent to {email_id}")
                update_logs(log_filename, email_id, f"Email {i+1} Sent")

                # Use the selected time interval
                st.info(f"Waiting for {time_intervals[i]} seconds...")
                for _ in range(time_intervals[i] // 5):  # Check every 5 seconds
                    time.sleep(5)
                    if check_for_replies(service, sent_message_id, email_id):
                        st.success(f"Reply received from {email_id}. Terminating pipeline.")
                        update_logs(log_filename, email_id, "Reply Received")
                        break
                else:
                    st.info(f"No reply after email {i+1}. Moving to next step.")

                if sent_message_id and check_for_replies(service, sent_message_id, email_id):
                    break
            except Exception as e:
                st.error(f"Error occurred during email pipeline for {email_id} - {company_name}: {e}")
                update_logs(log_filename, email_id, f"Error: {e}")
                break
        else:
            st.warning(f"No reply received after all emails to {email_id}. Pipeline completed.")
            update_logs(log_filename, email_id, "No Reply Received")

# Main Execution
def main():
    st.title("Email Automation Pipeline")

    # File upload for email template
    uploaded_file = st.file_uploader("Upload your Email Template (Word)", type=["docx"])
    email_templates = []  # Initialize as empty list to avoid UnboundLocalError

    if uploaded_file:
        email_templates = load_email_templates(uploaded_file)
        st.success("Email templates loaded successfully!")

    # Gmail Authentication
    if st.button("Authenticate Gmail"):
        try:
            service = authenticate_gmail()
            st.session_state.service = service  # Store the authenticated service in session state
            st.success("Gmail authenticated successfully!")
        except Exception as e:
            st.error(f"Error during Gmail authentication: {e}")
            return

    # Dynamic recipient input
    st.subheader("Enter Recipients")

    if "recipients" not in st.session_state:
        st.session_state.recipients = []  # Initialize the recipients list

    # Add recipient inputs
    with st.form("recipient_form"):
        email_id = st.text_input("Recipient Email ID", key="email_input")
        company_name = st.text_input("Company Name", key="company_input")
        add_recipient = st.form_submit_button("Add Recipient")
        if add_recipient and email_id and company_name:
            st.session_state.recipients.append({"email_id": email_id, "company_name": company_name})
            st.success(f"Added: {email_id} - {company_name}")

    # Display the recipients
    if st.session_state.recipients:
        st.subheader("Recipients List")
        for idx, recipient in enumerate(st.session_state.recipients):
            st.text(f"{idx + 1}. {recipient['email_id']} - {recipient['company_name']}")

    # Dynamic time interval selection
    st.subheader("Select Timer Interval (in seconds)")
    time_interval = st.slider("Choose the time interval for email sending:", min_value=10, max_value=300, value=60, step=10)

    # Start the email pipeline
    if st.button("Start Email Pipeline"):
        if "service" not in st.session_state:
            st.error("Please authenticate Gmail first.")
        elif not email_templates:
            st.error("Please upload an email template to proceed.")
        elif st.session_state.recipients:
            try:
                # Generate a unique log file name for each session
                log_filename = generate_log_filename()

                # Use the selected time interval
                time_intervals = [time_interval] * len(email_templates)

                # Pass `time_intervals` to the `email_pipeline` function
                email_pipeline(st.session_state.service, st.session_state.recipients, email_templates, log_filename, time_intervals)

                # After the email pipeline terminates, give option to download the log file
                st.success("Email pipeline completed. You can download the log file.")
                with open(log_filename, "rb") as file:
                    st.download_button("Download Email Log", file, file_name=log_filename)
            except Exception as e:
                st.error(f"Error during email pipeline execution: {e}")
        else:
            st.warning("No recipients added. Please add recipients to proceed.")


if __name__ == "__main__":
    main()
