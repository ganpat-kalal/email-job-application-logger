import os
import base64
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import openai

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
SCOPES_SPREADSHEET = ['https://www.googleapis.com/auth/spreadsheets']

# Authenticate with Gmail API
def authenticate_gmail():
    creds = None
    # The file token_gmail.json stores the user's access and refresh tokens
    if os.path.exists('token_gmail.json'):
        creds = Credentials.from_authorized_user_file('token_gmail.json')
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token_gmail.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def authenticate_google_sheets():
    creds = None
    # The file token_sheets.json stores the user's access and refresh tokens
    if os.path.exists('token_sheets.json'):
        creds = Credentials.from_authorized_user_file('token_sheets.json')
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES_SPREADSHEET)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token_sheets.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def extract_job_application_emails(service):
    results = service.users().messages().list(userId='me', maxResults=1).execute()  # Adjust maxResults as needed
    messages = results.get('messages', [])
    job_application_emails = []

    for message in messages:
        msg_id = message['id']
        msg = service.users().messages().get(userId='me', id=msg_id).execute()
        
        # Extracting date and time
        headers = msg['payload']['headers']
        date_header = next((header for header in headers if header['name'] == 'Date'), None)
        date_time = date_header['value'] if date_header else None

        # Extracting subject
        subject_header = next((header for header in headers if header['name'] == 'Subject'), None)
        subject = subject_header['value'] if subject_header else None

        # Extracting email content
        parts = msg['payload'].get('parts', [])
        if parts:
            body_data = parts[0]['body']['data']
            email_content = base64.urlsafe_b64decode(body_data).decode('utf-8')
        else:
            email_content = "No content found."

        # Append extracted information to the list
        job_application_emails.append({
            'id': msg_id,
            'date_time': date_time,
            'subject': subject,
            'content': email_content
        })

    return job_application_emails

def process_emails_with_chatgpt(emails):
    # Initialize OpenAI API
    openai.api_key = 'YOUR_OPENAI_API_KEY'
    processed_data = []
    for email in emails:
        # Extract email content
        email_data = email['payload']['parts'][0]['body']['data']
        email_content = base64.urlsafe_b64decode(email_data).decode('utf-8')
        # Process email content using ChatGPT API
        response = openai.Completion.create(
            engine="text-davinci-002",
            prompt=email_content,
            max_tokens=50
        )
        processed_data.append(response['choices'][0]['text'])
    return processed_data

def save_to_google_spreadsheet(data):
    creds = authenticate_google_sheets()
    service = build('sheets', 'v4', credentials=creds)
    
    values = [
        ['Email ID', 'Date & Time', 'Subject', 'Content']
    ]
    for email in data:
        values.append([
            email['id'],
            email['date_time'],
            email['subject'],
            email['content']
        ])

    service.spreadsheets().values().update(
        spreadsheetId='1Sh2cTuyH7rwazM9L7ebUbIijXA7rcdhoZ3dsHNo0axE',
        range='Sheet1!A1:D{}'.format(len(data) + 1),
        valueInputOption='RAW',
        body={'values': values}
    ).execute()

def main():
    # Authenticate with Gmail API
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)
    # Extract job application-related emails
    job_application_emails = extract_job_application_emails(service)
    # print(job_application_emails)
    # Process emails using ChatGPT API
    # processed_data = process_emails_with_chatgpt(job_application_emails)
    # Save processed data to Google Spreadsheet
    # save_to_google_spreadsheet(processed_data)
    save_to_google_spreadsheet(job_application_emails)

if __name__ == '__main__':
    main()
