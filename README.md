# Automated Email Sender

## Project Overview

The **Automated Email Sender** is a Python-based application developed using the Tkinter library for the GUI and the `smtplib` library for sending emails. The tool automates the process of sending personalized emails to a list of recipients imported from a CSV or Excel file. This project was developed as part of an internship at TECHPLEMENT.

## Features

- **SMTP Configuration:** Allows users to configure the SMTP server and port for sending emails. Supports secure connections via SSL/TLS.
- **Personalized Email Templates:** Users can create email templates with placeholders for personalization, such as recipient name and company.
- **Recipient Import:** Recipients can be imported from CSV or Excel files, with verification that necessary columns (Email, Name, Company) are present.
- **File Attachments:** Users can attach files to be sent along with the email.
- **Email Preview:** Allows users to preview the email before sending it to ensure the content is correctly formatted.
- **CAPTCHA Verification:** A simple CAPTCHA mechanism is implemented to prevent unauthorized sending of emails.
- **Status Tracking:** Real-time tracking of email sending status, informing users of successful sends or any errors encountered.
- **Error Handling:** Robust error handling ensures that issues like failed logins, incorrect recipient files, and email sending errors are gracefully managed.
- **Security:** Ensures secure handling of user credentials and implements measures like CAPTCHA to prevent misuse.
- **Performance Optimization:** Efficiently handles large recipient lists and prevents the program from being flagged as spam by email servers.

## How to Run the Application

1. **Install Dependencies:**
   - Ensure you have Python installed.
   - Install the required packages using pip:
     ```bash
     pip install pandas
     ```

2. **Run the Application:**
   - Execute the Python script to start the GUI application.
   - The Tkinter window will appear, allowing you to configure the email settings, import recipients, and send emails.

3. **SMTP Configuration:**
   - Enter the SMTP server details (e.g., `smtp.gmail.com`) and port (`587` for TLS).
   - Provide the sender's email address and password for authentication.

4. **Email Composition:**
   - Fill in the subject and body of the email. Use placeholders like `{name}` and `{company}` for personalization.
   - Attach a file if necessary.

5. **Recipient Import:**
   - Click the "Import CSV/Excel" button to load a recipient list.
   - Ensure the file contains the required columns: Email, Name, and Company.

6. **CAPTCHA Verification:**
   - Enter the CAPTCHA displayed in the interface to enable the "Send Email" button.

7. **Preview and Send:**
   - Preview the email to ensure it is correctly formatted.
   - Click "Send Email" to begin the process.
   - Monitor the status display to check which emails were successfully sent and which encountered errors.

## Requirements

- **Python 3.7+**
- **Libraries:**
  - `tkinter` for the GUI
  - `smtplib` for sending emails
  - `pandas` for handling CSV/Excel files
  - `ssl` for secure connections

## File Structure

```plaintext
Automated Email Sender/
│
├── main.py           # Main script to run the application
└── README.md         # Project documentation
└── requirements.txt  # List of dependencies
```

## Author

- **Darshan Kudache**
- **Tejushree R**

## License
This project is licensed under the MIT License.