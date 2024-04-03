# Job Application Tracking Automation

This project automates the process of tracking job application emails from Gmail and storing relevant information in a Google Spreadsheet.

## Prerequisites

Before running the application, ensure you have the following installed:

- Python 3.x
- Git

## Setup Instructions

1. **Clone the repository**:  
   ```bash
   git clone https://github.com/yourusername/job-application-tracking-automation.git
   ```
   
2. **Navigate to the project directory**:  
   ```bash
   cd job-application-tracking-automation
   ```

3. **Install the required Python packages**:  
   ```bash
   pip install -r requirements.txt
   ```

4. **Obtain credentials for Gmail API and Google Sheets API**:
   
   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Create a new project.
   - Enable the Gmail API and Google Sheets API for your project.
   - Create credentials (OAuth 2.0 client ID) for both APIs.
   - Download the credentials JSON file for each API and save them as `credentials_gmail.json` and `credentials_sheets.json` respectively in the project directory.

5. **Run the application**:  
   ```bash
   python JobApplicationTrackingSheet.py
   ```

6. **Follow the authentication flow to grant access to Gmail and Google Sheets**.

## Usage

Once the application is running, it will:

- Retrieve job application-related emails from your Gmail inbox.
- Extract the date, time, subject, and content of each email.
- Process the emails by using ChatGPT API 
- Store this information in a Google Spreadsheet

## Customization

You can customize the application by adjusting parameters such as the number of emails to retrieve (`maxResults`) and the spreadsheet ID and range in the `save_to_google_spreadsheet` function.

## Contributions

Contributions are welcome! If you have any suggestions or want to contribute to this project, feel free to open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
