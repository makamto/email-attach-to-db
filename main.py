import os
import imaplib
import email
from datetime import datetime, date, timedelta
import warnings
import pandas as pd
from sqlalchemy import create_engine, text
import psycopg2
from psycopg2 import Error

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl.*')

# Database connection function
def db_connection(DB_DATABASE, DB_USERNAME, DB_PASSWORD, DB_HOST, DB_PORT):
    try:
        connection = psycopg2.connect(
            user=DB_USERNAME,
            password=DB_PASSWORD,
            host=DB_HOST,
            port=DB_PORT,
            database=DB_DATABASE
        )
        connection.reset()
        cursor = connection.cursor()
        cursor.execute("SELECT version();")
        if connection:
            print("--- Connection is started ---\n")
        return connection
    except (Exception, Error) as error:
        print("--- Error while connecting to database ---\n", error)
        return None

# Check table columns function
def check_table_columns(engine, TABLE_NAME):
    with engine.connect() as conn:
        query = text(f"SELECT column_name FROM information_schema.columns WHERE TABLE_NAME = '{
                     TABLE_NAME}'")
        result = conn.execute(query)
        columns = [row[0] for row in result]
        return columns

# Database disconnection function
def db_disconnection(connection):
    if connection:
        connection.close()
        print("\n--- Connection is closed. ---")

# Data deduplication function
def deduplicate_data(df, engine, TABLE_NAME, UNIQUE_COLUMNS):
    try:
        df = df.astype({col: 'object' for col in UNIQUE_COLUMNS})
        columns = check_table_columns(engine, TABLE_NAME)
        if not all(col in columns for col in UNIQUE_COLUMNS):
            missing_columns = [
                col for col in UNIQUE_COLUMNS if col not in columns]
            raise ValueError(f"Missing columns in the table: {
                             missing_columns}")

        where_conditions = ' AND '.join(
            [f'"{col}" = :{col.replace(" ", "_")}' for col in UNIQUE_COLUMNS])
        query = text(f'SELECT * FROM "{TABLE_NAME}" WHERE {where_conditions}')
        params = {col.replace(" ", "_"): df[col].iloc[0]
                  for col in UNIQUE_COLUMNS}
        existing_rows = pd.read_sql_query(query, engine, params=params)
        duplicates = pd.merge(
            df, existing_rows, on=UNIQUE_COLUMNS, how='inner')
        deduplicated_df = df[~df.index.isin(duplicates.index)]
        return deduplicated_df
    except:
        print("You uploaded it before.")
        return pd.DataFrame()

# Function to upload Excel files to PostgreSQL
def upload_xlsx_to_postgresql(directory_path, TABLE_NAME, host, port, database, username, password, UNIQUE_COLUMNS, attachment_type, database_type):
    try:
        if not os.path.exists(directory_path):
            os.makedirs(directory_path)
            print(f"Directory '{directory_path}' created")

        files = os.listdir(directory_path)
        for file_name in files:
            if file_name.endswith(attachment_type):
                file_path = os.path.join(directory_path, file_name)
                if database_type.lower() == "postgresql":
                    db_url = f'postgresql://{username}:{password}@{host}:{port}/{database}'
                elif database_type.lower() == "mysql":
                    db_url = f'mysql://{username}:{password}@{host}:{port}/{database}'
                elif database_type.lower() == "sqlite":
                    db_url = f'sqlite:///{database}'
                elif database_type.lower() == "oracle":
                    db_url = f'oracle://{username}:{password}@{host}:{port}/'
                elif database_type.lower() == "mssql":
                    db_url = f'mssql+pymssql://{username}:{password}@{host}:{port}/{database}'
                else:
                    print("Unsupported database type. Please try again.")
                    exit(1)
                engine = create_engine(db_url)
                engine.expire_on_commit = False
                df = pd.read_excel(file_path)
                df_deduplicated = deduplicate_data(
                    df, engine, TABLE_NAME, UNIQUE_COLUMNS)
                if not df_deduplicated.empty:
                    df_deduplicated.to_sql(
                        TABLE_NAME, engine, if_exists='append', index=False)
                    print(f"Data from file '{file_name}' uploaded successfully to table '{
                          TABLE_NAME}' in the database.")
                    os.remove(file_path)
                    print(f"File '{file_name}' is deleted.")
                else:
                    print(f"No new data to upload from file '{file_name}'.")
                    os.remove(file_path)
                    print(f"File '{file_name}' is deleted.")
            else:
                print(f"File '{file_name}' does not exists")
    except Exception as e:
        print(f"An error occurred: {e}")

# Function to decode email headers
def decode_str(s):
    try:
        data = email.header.decode_header(s)
    except:
        return None
    sub_bytes = data[0][0]
    sub_charset = data[0][1]
    if sub_charset is None:
        return sub_bytes
    return sub_bytes.decode(sub_charset)

# Function to encode email subject
def encode_subject(subject):
    return email.header.Header(subject, 'utf-8').encode()

# Function to convert date to IMAP format
def imap_format_date(date):
    return date.strftime('%d-%b-%Y')

# Function to get date range
def get_date_input(prompt):
    while True:
        date_str = input(prompt)
        try:
            # Attempt to parse the date string into a data object
            date_obj = datetime.strptime(date_str, "%Y%m%d").date()
            return date_obj
        except ValueError:
            print("Invalid date format. Please enter the date in YYYYMMDD format.")

# Main function
if __name__ == "__main__":
    # Email server settings
    print("\n--- Your email setting: ---")
    EMAIL_HOST = input("Enter your host: ")
    EMAIL_PORT = input("Enter your port: ")  # or 143 for non-SSL
    EMAIL_USERNAME = input("Enter your email: ")
    EMAIL_PASSWORD = input("Enter your password: ")
    EMAIL_MAILBOX = input("Enter your mailbox (e.g, Inbox): ")
    EMAIL_SUBJECT = input("Enter your email subject (e.g, search subject starts with 'Potential'): ")  # can be fuzzy (Potential...)
    EMAIL_SUBJECT = EMAIL_SUBJECT or f'Re: {EMAIL_SUBJECT}'
    ATTACHMENT_TYPE = input("Enter your attachment type (e.g, .pdf, .xlsx, .docx...): ")
    SAVE_DIR = os.path.expanduser('~/Documents/attach-dir')


    # Database settings
    print("\n--- Your database setting: ---")
    DB_TYPE = input("Enter your database type (e.g. PostgreSQL, MySQL, SQLite...): ")
    DB_HOST = input("Enter your database host: ")
    DB_PORT = input("Enter your database port: ")
    DB_DATABASE = input("Enter your database name: ")
    DB_USERNAME = input("Enter your database username: ")
    DB_PASSWORD = input("Enter your database password: ")
    TABLE_NAME = input("Enter your table name: ")
    UNIQUE_COLUMNS = input("Enter your primary keys (comma-separated): ")
    UNIQUE_COLUMNS = [column.strip() for column in UNIQUE_COLUMNS.split(",")]

    print("\nYour attachment will automatically downloaded in '~/Documents/attach-dir'")

    # Date range settings
    yesterday = date.today() - timedelta(days=1)
    today = date.today()

    # Propmt for daily report or customize report
    print("\n--- Report Options: ---")
    print("1. Daily Report (Yesterday to Today)")
    print("2. Custom Report (Specify Dates)\n")
    option = input("Enter your choice (1/2): ")

    if option == "1":
        start_date = yesterday
        end_date = today
        print(f"\nDaily Report: From {start_date} to {end_date}: Searching...\n")
    else: 
        print("\n--- Custom Report: ---")
        start_date = get_date_input("Enter the start date (YYYYMMDD): ")
        end_date = get_date_input("Enter the end date (YYYYMMDD): ")
        print(f"From {start_date} to {end_date}: Searching email starts with \"{EMAIL_SUBJECT}...\"\n")

    # Format the dates for IMAP
    date_from = imap_format_date(start_date)
    date_to = imap_format_date(end_date)

    try:
        # Connect to the email server
        mail = imaplib.IMAP4_SSL(EMAIL_HOST, EMAIL_PORT)
        mail.login(EMAIL_USERNAME, EMAIL_PASSWORD)
        mail.select(EMAIL_MAILBOX)  # select the Inbox folder

        # Encode the email subject for the search query
        encoded_subject = encode_subject(EMAIL_SUBJECT)
        search_criteria = f'(SUBJECT "{
            EMAIL_SUBJECT}*" SINCE {date_from} BEFORE {date_to})'

        # Search for the email with the attachment
        status, response = mail.search(None, search_criteria)
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
            if subject.startswith(EMAIL_SUBJECT):
                # Loop through each part of the email
                for part in email_message.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    filename = decode_str(part.get_filename())

                    if 'attachment' in content_disposition and ATTACHMENT_TYPE in content_type or bool(filename):
                        if not filename:
                            filename = 'unknown'
                        attachment = part.get_payload(decode=True)

                        # Save the attachment to a file
                        if filename.endswith(ATTACHMENT_TYPE):
                            filepath = os.path.join(SAVE_DIR, filename)
                            with open(filepath, 'wb') as f:
                                f.write(attachment)

                            print(f'\n--- Attachment {filename} downloaded successfully! ---\n')

    except imaplib.IMAP4.error as e:
        print(f'Error: {e}')

    finally:
        # Close the email connection
        try:
            mail.close()
            mail.logout()
        except:
            pass
    
    print()

    # Connect to the database and upload data
    connection = db_connection(
        DB_DATABASE, DB_USERNAME, DB_PASSWORD, DB_HOST, DB_PORT)
    if connection:
        upload_xlsx_to_postgresql(SAVE_DIR, TABLE_NAME, DB_HOST, DB_PORT,
                                  DB_DATABASE, DB_USERNAME, DB_PASSWORD, UNIQUE_COLUMNS, ATTACHMENT_TYPE, DB_TYPE)
        db_disconnection(connection)
