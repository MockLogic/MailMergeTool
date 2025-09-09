Mail Merge Tool for Microsoft Outlook
=====================================

Version: 3.6 (2025-09-05)
Author: AI Assistant (Purple) / Brian Danford (Danford@uw.edu)

A robust Python script for creating personalized Outlook email drafts using CSV contact list and Markdown templates.

📦 Features:
- Automatic encoding detection for CSV files
- Smart handling of special characters in subjects/body
- Empty row skipping in CSV processing
- Markdown-to-HTML conversion with Outlook formatting
- Automatic attachment handling
- Comprehensive logging and error handling
- Windows-1252/UTF-8 CSV support
- Symbol cleaning for problematic characters

⚙️ Requirements:
- Windows OS with Microsoft Outlook installed
- Python 3.9 or newer
- Required Python packages:
  * pywin32 (for Outlook integration)
  * markdown (for template processing)
  * chardet (for encoding detection)

📥 Installation:
1. Install Python from python.org (check "Add to PATH" during installation)
2. Open Command Prompt and run:
   pip install pywin32 markdown chardet
	or
   python -m pip install pywin32 markdown chardet

📂 File Structure:
MailMergeTool/
├── MailMerge.py        # Main script
├── contacts.csv        # Recipient data (required)
├── email_template.md   # Email template markdown file (required)
├── Attachments/        # Folder for email attachments
├── run.bat             # Run this file to activate the script
└── ReadMe.txt          # This file

📝 Setup Instructions:

1. CSV File Preparation:
   - Required columns: To, CC, BCC, Subject, Attachments
   - Other columns are optional, and can be used in the e-mail template for insertion, a contents of a column like Name replace <<Name>>

2. Template File:
   - Create email_template.md using Markdown syntax (https://daringfireball.net/projects/markdown/syntax)
   - Use <<ColumnName>> placeholders for CSV data
   - Example:
     
     Hello <<Name>>,
     
     Please find attached <<DocumentName>>.
     

3. Attachments (optional):
   - Place files in the Attachments folder
   - Reference in CSV as comma-separated filenames in the attachment column.

🚀 Usage:
1. Place all files in the same directory
2. Open Outlook
2. Double click to execute run.bat

📋 Output:
- Creates Outlook drafts in your default Outlook profile
- Generates log file: mail_merge_YYYY-MM-DD_HH-MM-SS.log
- Success/failure summary in console and log

🔧 Troubleshooting:
Q: Getting encoding errors?
A: Ensure CSV is saved as UTF-8 or Windows-1252 encoding

Q: Empty rows being skipped?
A: This is intentional - check your CSV for blank lines

Q: Attachments not loading?
A: Verify:
   - Files are in Attachments folder
   - Filenames match exactly in CSV with commas between file names
   - No spaces before/after filenames in CSV

Q: Outlook not responding?
A: Check if Outlook is running in the background first

📄 License: Free for personal and business use. Modify as needed.

⚠️ Disclaimer:
- Always test with small batches first
- Verify drafts before sending
- Maintain data privacy - this script does NOT send emails automatically

📧 Support: Contact your local Python developer for assistance