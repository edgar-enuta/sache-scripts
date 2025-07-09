# Email to Excel Automation

## Introduction
This tool automatically reads specific emails, extracts important information, and saves it to an Excel file. It is designed for users who want to automate the process of collecting data from emails and organizing it in a spreadsheet, with no programming required.

## Installation
1. **Download and Install Python**
   - Download Python from [python.org](https://www.python.org/downloads/).
   - Install Python, making sure to check the box that says "Add Python to PATH" during installation.

2. **Download the Project**
   - Download or copy all the files and folders in this project to a folder on your computer.

3. **Install Required Packages**
   - Open the Command Prompt (press `Win + R`, type `cmd`, and press Enter).
   - Navigate to your project folder using the `cd` command. For example:
     ```
     cd D:\work\sache-scripts
     ```
   - Install the required packages by running:
     ```
     pip install -r requirements.txt
     ```

## Environment Configuration
1. **Create a `.env` File**
   - In your project folder, create a file named `.env` (if it does not already exist).

2. **Add the Following Settings to `.env`:**
   - Replace the values with your actual information.
     ```
     EMAIL=your_email@example.com
     EMAIL_PASSWORD=your_email_password
     IMAP_SERVER=imap.yourmailserver.com
     IMAP_PORT=993
     EXCEL_PATH=output/data.xlsx
     ```
   - `EMAIL`: Your email address.
   - `EMAIL_PASSWORD`: Your email password (or app password if using Gmail/Outlook).
   - `IMAP_SERVER`: The IMAP server address for your email provider.
   - `IMAP_PORT`: Usually 993 for secure connections.
   - `EXCEL_PATH`: The path where the Excel file will be saved (e.g., `output/data.xlsx`).

## Usage
1. **Start the Script**
   - In the Command Prompt, make sure you are in your project folder.
   - Run the following command:
     ```
     python main.py
     ```
   - The script will:
     - Connect to your email account.
     - Read unread emails that match the configured pattern.
     - Extract the required fields.
     - Save the data to the Excel file.
     - Mark the emails as processed.

2. **Check the Output**
   - Open the Excel file specified in your `.env` file to see the extracted data.

## How It Works
- The script connects to your email account using IMAP and searches for unread and unflagged emails.
- It only processes emails whose subject matches the pattern defined in `field_config.yaml` under `email_title.pattern`.
- For each matching email, it extracts fields based on the patterns and column names defined in `field_config.yaml`.
- The extracted data is appended to the Excel file specified by `EXCEL_PATH`.
- After successful export, the script flags the processed emails so they are not processed again.
- If any error occurs during the process, the script restores the Excel file to its previous state and removes the flag from any emails that were flagged during the failed run.

## Configuration Details
- **field_config.yaml**: This file controls which fields are extracted from emails. Each field has a `pattern` (regular expression) and an `excel_column` (the column name in Excel).
- **main.py**: This is the entry point. It handles the transaction logic, so no partial data is left if something fails.
- **imap.py**: Handles connecting to your email, searching, and fetching emails. It uses the IMAP protocol and does not mark emails as read when fetching.
- **extract_fields.py**: Extracts data from emails using the patterns in `field_config.yaml`.
- **excel.py**: Handles writing to the Excel file and backing up/restoring it for safe transactions.

## Safety and Data Integrity
- The script is transactional: if any step fails, all changes to the Excel file and email flags are rolled back.
- Emails are only flagged as processed after a successful export to Excel.
- The Excel file is backed up before writing and restored if an error occurs.

## Troubleshooting
- If you see errors about missing packages, make sure you have run `pip install -r requirements.txt`.
- If you have problems connecting to your email, double-check your `.env` settings.
- For Gmail, you may need to use an "App Password" and enable IMAP in your account settings.

## Notes
- Only emails that match the subject pattern in `field_config.yaml` are processed.
- The script uses the `\Flagged` IMAP flag to mark emails as processed. If you re-run the script, already flagged emails will not be processed again.
- Make sure your email provider supports IMAP and allows access from third-party apps.

## Changing the Email Pattern

If you want the script to process different types of emails, you can change the pattern it uses to recognize them.

1. Open the `field_config.yaml` file in a text editor.
2. Find the section called `email_title` and look for the line starting with `pattern:`. This is the rule the script uses to decide which emails to process.
3. To change which emails are processed, update the text after `pattern:`. For example, to match emails with a different subject, change the words or numbers in the pattern.

**Tip:**
- The pattern uses something called a "regular expression" (regex) to match email subjects. If you are not familiar with regex, you can use the website [regexr.com](https://regexr.com/) to test and understand patterns. Just copy your pattern there and try it out with example text.
- If you need to match a specific phrase, just put it in the pattern. For example:
  - To match subjects that start with "Invoice":
    ```yaml
    pattern: '^Invoice'
    ```
  - To match subjects that contain a 10-digit number:
    ```yaml
    pattern: '\\d{10}'
    ```

If you are unsure how to write your pattern, ask for help or use [regexr.com](https://regexr.com/) to experiment until it matches the emails you want.
