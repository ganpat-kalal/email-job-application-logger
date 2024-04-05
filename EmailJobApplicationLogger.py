import os
import base64
from datetime import datetime
from email.utils import parseaddr
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from openai import OpenAI
import re
import os

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
SCOPES_SPREADSHEET = ["https://www.googleapis.com/auth/spreadsheets"]


# Authenticate with Gmail API
def authenticate_gmail():
    creds = None
    # The file token_gmail.json stores the user's access and refresh tokens
    if os.path.exists("token_gmail.json"):
        creds = Credentials.from_authorized_user_file("token_gmail.json")
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token_gmail.json", "w") as token:
            token.write(creds.to_json())
    return creds


def authenticate_google_sheets():
    creds = None
    # The file token_sheets.json stores the user's access and refresh tokens
    if os.path.exists("token_sheets.json"):
        creds = Credentials.from_authorized_user_file("token_sheets.json")
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES_SPREADSHEET
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token_sheets.json", "w") as token:
            token.write(creds.to_json())
    return creds


def extract_job_application_emails(service):
    # query = "after:2024/01/01" # Getting only 100 emails (q=query)
    query = "(label:inbox OR label:sent)"
    results = (
        service.users().messages().list(userId="me", maxResults=10, q=query).execute()
    )  # Adjust maxResults as needed
    messages = results.get("messages", [])
    job_application_emails = []

    keywords = ["application", "Application", "bewerbung", "Bewerbung"]

    i = 1

    for message in messages:
        msg_id = message["id"]
        msg = service.users().messages().get(userId="me", id=msg_id).execute()

        # Extracting date and time
        headers = msg["payload"]["headers"]
        date_header = next(
            (header for header in headers if header["name"] == "Date"), None
        )
        if date_header:
            raw_date_time = date_header["value"]
            date_time_pattern = (
                r"(\w{3}, \d{1,2} \w{3} \d{4} \d{2}:\d{2}:\d{2})(?: \(\w+\))? \+\d{4}"
            )
            match = re.search(date_time_pattern, raw_date_time)
            if match:
                extracted_date_time = match.group(1)
                parsed_date_time = datetime.strptime(
                    extracted_date_time, "%a, %d %b %Y %H:%M:%S"
                )  # Parse raw date string
                formatted_date_time = parsed_date_time.strftime(
                    "%d-%m-%Y %H:%M:%S"
                )  # Format date and time
            else:
                formatted_date_time = None

        # Extracting sender's name and email
        sender_header = next(
            (header for header in headers if header["name"] == "From"), None
        )
        if sender_header:
            sender_name, sender_email = parseaddr(sender_header["value"])
        else:
            sender_name, sender_email = None, None

        # Extracting subject
        subject_header = next(
            (header for header in headers if header["name"] == "Subject"), None
        )
        subject = subject_header["value"] if subject_header else None

        # Extracting email content
        parts = msg["payload"].get("parts", [])
        email_content = ""
        # if parts:
        if parts:
            # body_data = parts[0]['body']['data']
            for part in parts:
                if part["mimeType"] == "text/plain":
                    # email_content = base64.urlsafe_b64decode(body_data).decode('utf-8')
                    email_content += base64.urlsafe_b64decode(
                        part["body"]["data"]
                    ).decode("utf-8")
            # else:
        else:
            if "snippet" in msg:
                #     email_content = "No content found."
                email_content = msg["snippet"]

        # Check if subject or email body contains any of the specified keywords
        if any(keyword in subject for keyword in keywords) or any(
            keyword in email_content for keyword in keywords
        ):

            # Append extracted information to the list
            job_application_emails.append(
                {
                    "id": i,
                    "date_time": formatted_date_time,
                    "sender_name": sender_name,
                    "sender_email": sender_email,
                    "subject": subject,
                    "content": email_content,
                }
            )

            i += 1

    return job_application_emails


def process_emails_with_chatgpt(emails):
    # email_content = emails[0]['content']
    # Initialize OpenAI API
    client = OpenAI(
        # This is the default and can be omitted
        api_key="YOUR_API_KEY",
    )
    processed_data = []
    for email in emails:
        # Extract email content
        # email_content = base64.urlsafe_b64decode(email["content"]).decode("utf-8")
        # Process email content using ChatGPT API
        response = client.chat.completions.create(
            messages=[
                {
                    "role": "system",
                    "content": "Check whether it is a job application email or not? If it is then extract only information company_name, Job_title and application_status. and I don't want any messages like 'Yes, this is a job application email.'",
                },
                {"role": "user", "content": email["content"]},
            ],
            model="gpt-3.5-turbo", 
            n=1,
        )
        processed_data.append(response.choices[0].message.content)
    print(processed_data)
    return processed_data


def parse_job_string(job_str):
    job_dict = {}
    for item in job_str.split("\n"):
        key_value = item.split(": ")
        if len(key_value) == 2:
            key, value = key_value
            job_dict[key.strip()] = value.strip()
        elif len(key_value) == 1:
            key_value = key_value[0].split("**")
            if len(key_value) == 2:
                key, value = key_value
                job_dict[key.strip()] = value.strip()
            else:
                # If unexpected format, skip this key-value pair
                continue
    return job_dict


def save_to_google_spreadsheet(data):
    creds = authenticate_google_sheets()
    service = build("sheets", "v4", credentials=creds)

    values = [
        ["No", "Date & Time", "Sender name", "Sender email", "Subject"]  # , 'Content'
    ]
    for email in data:
        values.append(
            [
                email["id"],
                email["date_time"],
                email["sender_name"],
                email["sender_email"],
                email["subject"],
                # email['content']
            ]
        )

    service.spreadsheets().values().update(
        spreadsheetId="1Sh2cTuyH7rwazM9L7ebUbIijXA7rcdhoZ3dsHNo0axE",
        range="Sheet1!A1:E{}".format(len(data) + 1),
        valueInputOption="RAW",
        body={"values": values},
    ).execute()


def main():
    # Authenticate with Gmail API
    creds = authenticate_gmail()
    service = build("gmail", "v1", credentials=creds)
    # Extract job application-related emails
    job_application_emails = extract_job_application_emails(service)
    # print(job_application_emails)
    # Process emails using ChatGPT API
    processed_data = process_emails_with_chatgpt(job_application_emails)
    # Save processed data to Google Spreadsheet
    # save_to_google_spreadsheet(processed_data)
    # save_to_google_spreadsheet(job_application_emails)


if __name__ == "__main__":
    main()
