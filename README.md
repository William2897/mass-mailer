# Bulk Mailer

A versatile bulk email automation tool that allows you to send personalized emails to multiple recipients using Outlook.

## Overview

Bulk Mailer is designed to streamline the process of sending personalized emails to a large number of recipients. The tool supports template-based emails with variable substitution, offering a user-friendly interface to manage your email campaigns.

## Features

- **Template Management**: Load and save email templates
- **CSV Integration**: Import recipient data from CSV files
- **Variable Substitution**: Personalize emails using variables from your CSV file
- **Outlook Integration**: Seamlessly works with your Outlook account
- **Attachment Support**: Add attachments to your emails
- **Signature Selection**: Choose from your existing Outlook signatures
- **Preview Functionality**: Preview emails before sending

## Versions

This repository contains multiple implementations of the Bulk Mailer tool:

- **Python GUI Version** (`bulk_mailer.py`): Standard Python implementation using Tkinter
- **Python GUI Enhanced** (`bulk_mailer v2.py`): Extended version with signature selection feature
- **VBA Version** (`bulk maler.bas`): Excel/VBA implementation for users who prefer working within Office

## Requirements

### Python Version
- Python 3.6+
- pandas
- tkinter
- pywin32 (win32com)
- beautifulsoup4

```
pip install pandas pywin32 beautifulsoup4
```

### VBA Version
- Microsoft Excel with VBA
- Microsoft Outlook

## Usage

### Python Version

1. Run the Python script:
   ```
   python bulk_mailer.py
   ```
   or for the enhanced version:
   ```
   python "bulk_mailer v2.py"
   ```

2. **Email Content**: Write or load your email template. Use `{ColumnName}` syntax for variables.

3. **Load Recipients**: Click "Load Recipients" to import a CSV file containing recipient information.

4. **Variable Insertion**: Select a variable from the dropdown and click "Insert" to add it to your template.

5. **Subject Line**: Enter your email subject (can include variables).

6. **Preview**: Click "Text Preview" to see how your email will look for the first recipient.

7. **Attachments**: Click "Add Attachment" to include files with your emails.

8. **Signature**: In v2, select your preferred Outlook signature.

9. **Send Emails**: Click "Send Emails" to send personalized emails to all recipients.

### VBA Version

1. Import the `bulk maler.bas` module into your Excel VBA project.

2. Create a UserForm named "BulkMailerForm" with the necessary controls.

3. Run the application from Excel and follow similar steps as the Python version.

## CSV File Format

Your CSV file should include the following columns:
- `Email`: Recipient's email address (required)
- `CC`: Carbon copy email addresses (optional)
- `BCC`: Blind carbon copy email addresses (optional)
- Any other columns will be available as variables for personalization

Example:
```
Email,Name,Company,CC,BCC
john@example.com,John Doe,ACME Inc.,manager@example.com,records@example.com
jane@example.com,Jane Smith,XYZ Corp.,,archive@example.com
```

## File Structure

- `bulk_mailer.py` - Main Python implementation
- `bulk_mailer v2.py` - Enhanced Python implementation with signature selection
- `bulk maler.bas` - VBA implementation
- `Recepients.csv` - Example/template CSV file
- Sample files for testing template loading/saving

## Notes

- The application requires Outlook to be installed and configured on your system.
- For large recipient lists, consider sending in batches to avoid triggering spam filters.
- Always preview your emails before sending to ensure proper formatting and variable substitution.
