import imaplib
import email
import os
import openpyxl

# Email server settings
HOST = ''
PORT = 993  # or 143 for non-SSL
USERNAME = ''
PASSWORD = ''
MAILBOX = ''
EMAIL_SUBJECT = ''
ATTACHMENT_TYPE = ''
SAVE_DIR = r''


def decode_str(s):  # Function to decode email headers
    try:
        data = email.header.decode_header(s)
    except:
        return None
    sub_bytes = data[0][0]
    sub_charset = data[0][1]
    if sub_charset is None:
        return sub_bytes
    return sub_bytes.decode(sub_charset)


def encode_subject(subject):  # Function to encode email subject
    return email.header.Header(subject, 'utf-8').encode()


try:
    # Connect to the email server
    mail = imaplib.IMAP4_SSL(HOST, PORT)
    mail.login(USERNAME, PASSWORD)
    mail.select(MAILBOX)  # select the Inbox folder

    # Encode the email subject for the search query
    encoded_subject = encode_subject(EMAIL_SUBJECT)
    search_criteria = f'(SUBJECT "{encoded_subject}")'

    # Search for the email with the attachment
    status, response = mail.search(None, '(SUBJECT EMAIL_SUBJECT)')
    if status != 'OK':
        print('No emails found!')
        exit()

    email_ids = response[0].split()

    # Create the save directory if it doesn't exist
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    # Loop through each email and download the attachment
    for num in email_ids:
        status, response = mail.fetch(num, '(RFC822)')
        if status != 'OK':
            print('Failed to fetch email.')
            continue

        raw_email = response[0][1]
        email_message = email.message_from_bytes(raw_email)

        # Decode the email subject
        subject = decode_str(email_message['Subject'])
        print(f'Email subject: {subject}')

        # Filter emails based on the subject
        if subject == EMAIL_SUBJECT:
            # Loop through each part of the email
            for part in email_message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                filename = decode_str(part.get_filename())

                if 'attachment' in content_disposition and ATTACHMENT_TYPE in content_type or bool(filename):
                    if not filename:
                        filename = 'unknown.xlsx'
                    attachment = part.get_payload(decode=True)

                    # Save the attachment to a file
                    if filename.endswith(ATTACHMENT_TYPE):

                        filepath = os.path.join(SAVE_DIR, filename)
                        with open(filepath, 'wb') as f:
                            f.write(attachment)

                        print(f'Attachment {
                              filename} downloaded successfully!')

except imaplib.IMAP4.error as e:
    print(f'Error: {e}')


finally:
    # Close the email connection
    try:
        mail.close()
        mail.logout()
    except:
        pass
