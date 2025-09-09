"""
OUTLOOK MAIL MERGE WORKING VERSION (FINAL FIXES)
Author: AI Assistant
Date: 2025-09-05
Version: 3.6
"""

# --------------------------
# IMPORTS & CONFIGURATION
# --------------------------
import sys
import csv
import re
import logging
import traceback
import time
from pathlib import Path
import markdown
from win32com.client import Dispatch
from chardet import detect

# Configuration
SCRIPT_DIR = Path(__file__).parent.resolve()
CSV_FILE = SCRIPT_DIR / 'contacts.csv'
TEMPLATE_FILE = SCRIPT_DIR / 'email_template.md'
ATTACHMENT_DIR = SCRIPT_DIR / 'Attachments'
LOG_FILE = SCRIPT_DIR / f'mail_merge_{time.strftime("%Y-%m-%d_%H-%M-%S")}.log'
REQUIRED_COLUMNS = ['To', 'CC', 'BCC', 'Subject', 'Attachments']

# Encoding detection parameters
ENCODING_CANDIDATES = ['utf-8-sig', 'cp1252', 'iso-8859-1']
MIN_CONFIDENCE = 0.7

# Symbol replacement mapping
CLEANING_MAP = {
    '\u2013': '-',  # En-dash
    '\u2014': '--', # Em-dash
    '\u2018': "'",  # Left single quote
    '\u2019': "'",  # Right single quote
    '\u201C': '"',  # Left double quote
    '\u201D': '"',  # Right double quote
    '\u00A0': ' ',  # Non-breaking space
    '\u2026': '...',# Ellipsis
    '\x96': '-'     # Windows-1252 en-dash
}

# --------------------------
# LOGGING CONFIGURATION (FIXED)
# --------------------------
def configure_logging():
    """Enhanced logging configuration with console+file output"""
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # Clear existing handlers
    if logger.hasHandlers():
        logger.handlers.clear()

    # File handler
    file_handler = logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_formatter)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('[%(levelname)s] %(message)s')
    console_handler.setFormatter(console_formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

# --------------------------
# ENCODING HANDLING (FIXED)
# --------------------------
def detect_encoding(file_path):
    """Detect file encoding with confidence checking"""
    with open(file_path, 'rb') as f:
        rawdata = f.read(10000)  # Read first 10KB for detection
        result = detect(rawdata)
    
    if result['confidence'] > MIN_CONFIDENCE:
        return result['encoding']
    return 'utf-8-sig'  # Fallback encoding

def safe_csv_reader(file_path):
    """Handle CSV file reading with encoding detection"""
    try:
        encoding = detect_encoding(file_path)
        logging.info(f"Detected encoding: {encoding}")
        return open(file_path, 'r', encoding=encoding, newline='')
    except UnicodeDecodeError:
        logging.warning("Detection failed, trying fallback encodings")
        for enc in ENCODING_CANDIDATES:
            try:
                f = open(file_path, 'r', encoding=enc, newline='')
                logging.info(f"Success with {enc}")
                return f
            except UnicodeDecodeError:
                continue
        raise ValueError("Failed to find suitable encoding")

# --------------------------
# VALIDATION FUNCTIONS (FIXED)
# --------------------------
def validate_environment():
    """Check required files and directories exist"""
    logging.info("Starting environment validation")
    
    missing = []
    if not CSV_FILE.exists():
        missing.append(f"CSV file: {CSV_FILE}")
    if not TEMPLATE_FILE.exists():
        missing.append(f"Template file: {TEMPLATE_FILE}")
    if not ATTACHMENT_DIR.exists():
        logging.warning(f"Attachment directory not found: {ATTACHMENT_DIR}")
    
    if missing:
        msg = "Missing required resources:\n" + "\n".join(missing)
        logging.critical(msg)
        raise FileNotFoundError(msg)

    logging.info("Environment validation passed")

def validate_csv_headers(reader):
    """Ensure CSV contains required columns"""
    missing = [col for col in REQUIRED_COLUMNS if col not in reader.fieldnames]
    if missing:
        msg = f"CSV missing required columns: {', '.join(missing)}"
        logging.critical(msg)
        raise ValueError(msg)
    
    logging.info("CSV header validation passed")

# --------------------------
# DATA PROCESSING (FIXED)
# --------------------------
def is_empty_row(row):
    """Check if all fields in a CSV row are empty"""
    return all(value.strip() == '' for value in row.values())

def clean_string(text):
    """Clean problematic characters from strings"""
    if not text:
        return text
        
    cleaned = text
    for char, replacement in CLEANING_MAP.items():
        cleaned = cleaned.replace(char, replacement)
    
    # Remove other non-ASCII characters
    cleaned = cleaned.encode('ascii', 'ignore').decode('ascii')
    return cleaned.strip()

def process_csv(file_path):
    """Load and validate CSV with enhanced error handling"""
    logging.info(f"Processing CSV: {file_path}")
    
    with safe_csv_reader(file_path) as f:
        reader = csv.DictReader(f)
        validate_csv_headers(reader)
        
        csv_data = []
        for idx, row in enumerate(reader, 2):  # Start at line 2 (1-based)
            if is_empty_row(row):
                logging.warning(f"Skipping empty row at line {idx}")
                continue
            
            # Clean all string fields
            cleaned_row = {k: clean_string(v) for k, v in row.items()}
            csv_data.append(cleaned_row)
    
    if not csv_data:
        logging.warning("CSV file contains no valid data after filtering")
        return None
    
    logging.info(f"Loaded {len(csv_data)} valid records")
    return csv_data

# --------------------------
# EMAIL CREATION (FIXED)
# --------------------------
def markdown_to_outlook(md_text, variables):
    """Convert Markdown to Outlook HTML with variable substitution"""
    try:
        substituted = 0
        for key, value in variables.items():
            placeholder = f'<<{key}>>'
            if placeholder in md_text:
                md_text = md_text.replace(placeholder, str(value))
                substituted += 1
        
        html = markdown.markdown(
            md_text,
            extensions=['extra', 'md_in_html'],
            output_format='html5'
        )
        
        return (
            '<div style="font-family: Calibri, sans-serif; font-size: 11pt; margin: 0;">'
            f'{html}'
            '</div>'
        )
    except Exception as e:
        logging.error(f"Markdown conversion failed: {str(e)}")
        raise

def create_draft(row, template, outlook):
    """Create Outlook draft with comprehensive error handling"""
    try:
        mail = outlook.CreateItem(0)
        mail.To = row.get('To', '').strip()
        mail.CC = row.get('CC', '').strip()
        mail.BCC = row.get('BCC', '').strip()
        mail.Subject = clean_string(row.get('Subject', 'No Subject'))
        
        mail.HTMLBody = markdown_to_outlook(template, row)
        
        if attachments := row.get('Attachments', ''):
            for filename in attachments.split(','):
                clean_name = filename.strip()
                if not clean_name:
                    continue
                
                path = ATTACHMENT_DIR / clean_name
                if path.exists():
                    mail.Attachments.Add(str(path))
                else:
                    logging.warning(f"Missing attachment: {clean_name}")
        
        mail.Save()
        return True
    except Exception as e:
        logging.error(f"Draft creation failed: {str(e)}")
        return False

# --------------------------
# MAIN EXECUTION (FIXED)
# --------------------------
def main():
    """Main control flow with error handling"""
    logging.info("=== Mail Merge Started ===")
    outlook = None
    
    try:
        validate_environment()
        
        logging.info(f"Loading template: {TEMPLATE_FILE}")
        template = TEMPLATE_FILE.read_text(encoding='utf-8')
        
        if (csv_data := process_csv(CSV_FILE)) is None:
            return
        
        outlook = Dispatch('Outlook.Application')
        success_count = 0
        
        for idx, row in enumerate(csv_data, 1):
            if create_draft(row, template, outlook):
                success_count += 1
        
        logging.info(f"Created {success_count}/{len(csv_data)} drafts")
        print(f"Successfully created {success_count} email drafts")
    
    except Exception as e:
        logging.critical(f"Fatal error: {str(e)}\n{traceback.format_exc()}")
        print(f"Error: {str(e)} (see {LOG_FILE} for details)")
    
    finally:
        if outlook:
            outlook.Quit()
        logging.info("=== Mail Merge Completed ===")

if __name__ == '__main__':
    configure_logging()
    main()