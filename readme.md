
Main file is credentials.json and email_agent.py


**# Email Automation Pipeline**

This project automates the process of sending emails to multiple recipients, checking for replies, and logging the results. The pipeline leverages Gmail API for sending and receiving emails, while utilizing a Word document as a source for email templates. The status of each email is logged in an Excel file, including information such as whether the email was sent, if a reply was received, and timestamps for each action.

## Features

- Send automated emails to multiple recipients.
- Customizable email templates loaded from a Word document.
- Track replies to each email.
- Log the status of each email (Sent, No Reply, Reply Received) with timestamps.
- User-friendly interface built with Streamlit for dynamic recipient management and email pipeline control.

## Requirements

run environment - 

.\env\Scripts\activate

Before running the application, you need to install the required Python packages. Use the following command:

```bash
pip install -r requirements.txt







## Setup

- Create a "credentials.json" file
- To authenticate the Gmail API, you'll need a credentials.json file. 

Follow the steps below to obtain the credentials:

- Visit the Google Developers Console.
- Create a new project and enable the Gmail API.
- Create OAuth 2.0 credentials and download the credentials.json file.
- Place this file in the root directory of your project.
-

## Prepare the Email Templates

- Create a Word document (.docx) that contains the email templates.

Each email template should have a structure like this:

Email_subject: <subject>
Email_body: <body>


You can customize the placeholders (e.g., {{company_name}}) in the subject and body, which will be dynamically replaced during execution.




## Usage

1. Start the Streamlit App
   Run the following command to launch the Streamlit app:

- streamlit run email_agent.py


2. Interact with the Interface

- Upload Email Template: Upload your Word document containing email templates.
- Authenticate Gmail: Click the "Authenticate Gmail" button to authenticate using your credentials.json.
- Add Recipients: Add recipients by entering their email addresses and associated company names.
- Start Email Pipeline: Click the "Start Email Pipeline" button to begin sending emails and checking for replies.


3. Log Files

The email log file Email Logging.xlsx will be created/updated in the project directory, where each action (sending emails, receiving replies) is logged with timestamps.

Example Workflow:

- Upload an email template.
- Authenticate with Gmail.
- Add recipients.
- Start the email pipeline to send the emails, wait for replies, and log all actions.


Troubleshooting

- Error: HttpError during Gmail API interaction:
  Ensure your Gmail API credentials are correctly configured and that the Gmail API is enabled in your Google Cloud project.

- Error: FileNotFoundError for log file:
  The log file will be created if it does not exist. Make sure the application has write permissions to the directory where the log file is stored.




sample email templates:::


"""
  
Email_subject: Collaboration Opportunity: ABC and {company_name}

Email_body: Dear {company_name} Team,

I hope this message finds you well! I'm John, Sales manager at ABC. I came across {company_name} and was impressed by your work in the AI/ML and Software development space.

At ABC, we specialize in AI Applications and Full-Stack Development. Given {company_name}'s focus on AI/ML and Software development, I believe there's potential for a meaningful partnership between our organizations.

Would it be possible to schedule a quick call to explore how ABC and {company_name} might work together? We're confident we can create value for both sides.

Looking forward to hearing from you!

Best regards,
John Doe
Sales manager
john@abc.com
+1-8765-434-568

"""