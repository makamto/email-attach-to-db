**Email Attachment Downloader**
=============================

**Description**
This Python script connects to an email server using IMAP, searches for emails with a specific subject, and downloads attachments of a specified type to a local directory.

**Configuration**
Before running the script, you need to configure the following variables at the top of the script:

* `HOST`: The hostname of your email server.
* `PORT`: The port number to use for the IMAP connection (default is 993 for SSL).
* `USERNAME`: Your email username.
* `PASSWORD`: Your email password.
* `MAILBOX`: The mailbox to search for emails (e.g. "INBOX").
* `EMAIL_SUBJECT`: The subject of the emails to search for.
* `ATTACHMENT_TYPE`: The file extension of the attachments to download (e.g. "xlsx").
* `SAVE_DIR`: The directory where attachments will be saved.

**Usage**
1. Update the configuration variables at the top of the script.
2. Run the script using Python (e.g. `python email_attachment_downloader.py`).
3. The script will connect to the email server, search for emails with the specified subject, and download attachments of the specified type to the specified directory.

**Note**
This script uses the `imaplib` and `email` libraries to connect to the email server and parse email messages. It also uses the `openpyxl` library to handle Excel file attachments. Make sure you have these libraries installed before running the script.

**Troubleshooting**
If you encounter any issues while running the script, check the error messages for more information. You can also try debugging the script by adding print statements or using a Python debugger.
