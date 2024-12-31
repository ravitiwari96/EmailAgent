# Email Automation Pipeline

This project automates the process of sending emails to multiple recipients, checking for replies, and logging the results. The pipeline leverages Gmail API for sending and receiving emails while utilizing a Word document as a source for email templates. The status of each email is logged in an Excel file, including information such as whether the email was sent, if a reply was received, and timestamps for each action.


## Features
- **Send Automated Emails:** Automate the process of sending emails to multiple recipients.
- **Customizable Email Templates:** Load and use email templates from a Word document with dynamic placeholders.
- **Track Replies:** Monitor and log replies to each email in real time.
- **Status Logging:** Record the status of each email (Sent, No Reply, Reply Received) with timestamps in an Excel log file.
- **Streamlit Interface:** Interact with a user-friendly Streamlit interface for managing recipients and controlling the email pipeline.



## Requirements
- **Run Environment:**  .\env\Scripts\activate

- **Before running the application, install the required Python packages:**

  pip install -r requirements.txt





  ## Setup
- **Create a "credentials.json" File:** To authenticate the Gmail API, you'll need a credentials.json file. Follow the steps below to obtain the credentials:
- Visit the Google Developers Console.
- Create a new project and enable the Gmail API.
- Create OAuth 2.0 credentials and download the credentials.json file.
- Place this file in the root directory of your project.
