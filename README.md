# Automated Excel Report Email Sender

## Project Overview

A Python automation tool that sends Excel reports directly through Microsoft Outlook and maintains a log of every email sent.

## Features

- Browse and select an Excel report
- Send reports through Microsoft Outlook
- Support single or multiple recipients
- Add custom email message and subject
- Attach Excel report automatically
- Generate email log in CSV format
- Error handling with message dialogs

## Technologies Used

- Python
- Tkinter
- PyWin32
- CSV
- OS
- Datetime

## Folder Structure

```text
Automated-Excel-Report-Email-Sender/
│
├── excel_email_sender.py
├── email_log.csv
├── requirements.txt
├── LICENSE
└── README.md
```

## Outlook Setup

This project uses Microsoft Outlook through PyWin32 to send Excel reports.

1. Install Microsoft Outlook on Windows.
2. Sign in to your Outlook account.
3. Make sure Outlook is configured and able to send emails.
4. Open Outlook at least once before running the Python script.
5. Install the required Python packages.
6. Run the Python script.

## Requirements

- Windows Operating System
- Microsoft Outlook Installed
- Outlook Account Configured
- Python 3.x
- PyWin32

## Output

After execution, the project:

- Allows the user to select an Excel report
- Collects recipient email addresses
- Accepts a custom subject and message
- Sends the report through Microsoft Outlook
- Records the email details in email_log.csv
- Displays a success or error message

## Future Improvements

- CC and BCC support
- Scheduled email sending
- Multiple attachments

## License

This project is licensed under the MIT License.

## Author

**Akshay Gawand**