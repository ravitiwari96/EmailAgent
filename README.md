# Email Automation 

This project automates the process of sending emails to multiple recipients, checking for replies, and logging the results. The pipeline leverages Gmail API for sending and receiving emails while utilizing a Word document as a source for email templates. The status of each email is logged in an Excel file, including information such as whether the email was sent, if a reply was received, and timestamps for each action.



## Features
- **Send Automated Emails:** Automate the process of sending emails to multiple recipients.
- **Customizable Email Templates:** Load and use email templates from a Word document with dynamic placeholders.
- **Track Replies:** Monitor and log replies to each email in real time.
- **Status Logging:** Record the status of each email (Sent, No Reply, Reply Received) with timestamps in an Excel log file.
- **Streamlit Interface:** Interact with a user-friendly Streamlit interface for managing recipients and controlling the email pipeline.



## Requirements
- **Run Environment:**  .\env\Scripts\activate

- **Before running the application, install the required Python packages:** pip install -r requirements.txt

## Setup
 **Create a "credentials.json" File:** To authenticate the Gmail API, you'll need a credentials.json file. Follow the steps below to obtain the credentials:
- Visit the Google Developers Console.
- Create a new project and enable the Gmail API.
- Create OAuth 2.0 credentials and download the credentials.json file.
- Place this file in the root directory of your project.

## Prepare the Email Templates
- Create a Word document (.docx) containing the email templates.
-  Structure each email template like this:

  **Email_subject: <subject>  
    Email_body: <body>**  

- Customize placeholders **(e.g., {{company_name}})** in the subject and body. These will be dynamically replaced during execution.

## Usage
1. **Start the Streamlit App**
- **Run the following command to launch the Streamlit app:** streamlit run email_agent.py  

2. **Interact with the Interface**
- **Upload Email Template:** Upload your Word document containing email templates.
- **Authenticate Gmail:** Click the "Authenticate Gmail" button to authenticate using your credentials.json.
- **Add Recipients:** Add recipients by entering their email addresses and associated company names.
- **Start Email Pipeline:** Click the "Start Email Pipeline" button to send emails and check for replies.

3. **Log Files**
- The email log file **(Email Logging.xlsx)** will be created/updated in the project directory. It logs all actions, including sending emails, receiving replies, and timestamps.

<img width="622" alt="image" src="https://github.com/user-attachments/assets/5f3be8e0-4ba2-43dd-9a39-d2083fef40a2" />



<img width="608" alt="image" src="https://github.com/user-attachments/assets/442ba704-554e-4b63-a5f5-39b139c5e80f" />



<img width="599" alt="image" src="https://github.com/user-attachments/assets/997ff203-ee17-442b-888d-22c4eb10a8f8" />



## Example Workflow
- Upload an email template.
- Authenticate with Gmail.
- Add recipients.
- Start the email pipeline to send emails, wait for replies, and log all actions.

## Troubleshooting
- **Error: HttpError During Gmail API Interaction:**
  Ensure your Gmail API credentials are correctly configured, and the Gmail API is enabled in your Google Cloud project.

- **Error: FileNotFoundError for Log File:**
  The log file will be created if it does not exist. Ensure the application has write permissions to the directory where the log file is stored.





