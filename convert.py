import sys
import mailbox
import os
import email
import win32com.client
import pythoncom
from tqdm import tqdm  # Progress bar
import logging
import concurrent.futures
import gc
import time
import hashlib
import argparse
import traceback
from retrying import retry
import signal
import shutil
import json
from io import BytesIO
import multiprocessing

# Signal Handling for Graceful Exit
def signal_handler(sig, frame):
    print("\nGracefully shutting down...")
    save_checkpoint()  # Ensure progress is saved before exit
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)

# Setup logging and command-line parsing
def setup_logging_and_parse_args():
    """Setup logging and parse command-line arguments."""
    parser = argparse.ArgumentParser(description="Convert MBOX to PST using Outlook")
    parser.add_argument("--mbox_file", help="Path to the MBOX file. If not provided, the script will auto-detect all '.mbox' files in the current directory.")
    parser.add_argument("output_folder", help="Folder to save the attachments")
    parser.add_argument("--pst_file", help="Path to the PST file", default="emails.pst")
    parser.add_argument("--log-level", help="Set the log level (DEBUG, INFO, WARNING, ERROR, CRITICAL)", default="INFO")
    parser.add_argument("--batch-size", help="Set the batch size for processing emails", type=int, default=500)
    parser.add_argument("--workers", help="Number of parallel workers", type=int, default=multiprocessing.cpu_count())
    return parser.parse_args()

# Initialize logging and processed email checkpoint
def initialize_logging_and_checkpoint(log_level):
    """Initialize logging and load the email checkpoint file."""
    log_level = getattr(logging, log_level.upper(), logging.INFO)
    logging.basicConfig(filename='import_log.txt', level=log_level,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    return load_checkpoint()

def load_checkpoint():
    """Load the set of processed emails from a checkpoint file."""
    if os.path.exists('processed_emails.json'):
        with open('processed_emails.json', 'r') as f:
            return set(json.load(f))
    return set()

def save_checkpoint():
    """Save the set of processed emails to a checkpoint file."""
    with open('processed_emails.json', 'w') as f:
        json.dump(list(processed_emails), f)

def find_mbox_files():
    """Find all .mbox files in the current directory."""
    print("Searching for .mbox files in the current directory...")
    mbox_files = [f for f in tqdm(os.listdir('.'), desc="Looking for MBOX files") if f.endswith('.mbox')]
    return mbox_files

def hash_email(raw_email):
    """Create a unique hash of the email content to track processed emails."""
    return hashlib.sha256(raw_email.encode('utf-8')).hexdigest()

def extract_emails_from_mbox_stream(mbox_file):
    """Stream emails from MBOX file instead of loading all at once."""
    try:
        print(f"Opening MBOX file: {mbox_file}")
        mbox = mailbox.mbox(mbox_file)
        print("Streaming emails from MBOX file...")
        for message in tqdm(mbox, desc="Streaming emails"):
            yield message.as_string()
    except Exception as e:
        logging.error(f"Error extracting emails from MBOX: {e}")
        raise

def save_attachment_in_memory(part, mail_item):
    """Save the attachment from the email part to the Outlook mail item using memory buffer."""
    filename = part.get_filename()
    if filename:
        payload = part.get_payload(decode=True)
        if payload:
            attachment_stream = BytesIO(payload)
            mail_item.Attachments.Add(attachment_stream, filename)

@retry(wait_exponential_multiplier=1000, wait_exponential_max=10000, stop_max_attempt_number=5)
def ensure_outlook_running():
    """Ensure that Outlook is running and accessible."""
    try:
        pythoncom.CoInitialize()  # Initialize COM library
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as e:
        logging.error(f"Failed to connect to Outlook: {e}")
        raise

def release_com_object(obj):
    """Properly release a COM object to avoid memory leaks."""
    if obj:
        obj = None

def process_email_with_retry(raw_email, inbox_folder, output_folder, retries=3):
    """Attempts to process an email and import it into Outlook."""
    attempt = 0
    email_id = hash_email(raw_email)
    if email_id in processed_emails:
        logging.info(f"Skipping already processed email: {email_id}")
        return
    
    while attempt < retries:
        try:
            process_email(raw_email, inbox_folder, output_folder)
            processed_emails.add(email_id)  # Mark email as processed
            return  # If successful, break out of the retry loop
        except Exception as e:
            attempt += 1
            logging.error(f"Error processing email on attempt {attempt}: {e}")
            time.sleep(1)  # Wait for 1 second before retrying
            if attempt >= retries:
                logging.error(f"Failed to process email after {retries} attempts: {e}")

def process_email(raw_email, inbox_folder, output_folder):
    """Processes an individual email and adds it to the Outlook PST."""
    try:
        logging.info(f"Processing email: {raw_email[:50]}...")
        msg = email.message_from_string(raw_email)
        mail_item = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
        mail_item.Subject = msg['subject'] or "(No Subject)"
        mail_item.To = msg.get('to', '')
        mail_item.CC = msg.get('cc', '')
        mail_item.BCC = msg.get('bcc', '')
        mail_item.SentOn = msg.get('date', '')
        mail_item.SenderEmailAddress = msg.get('from', '')

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if part.get_filename():
                    logging.info(f"Saving attachment: {part.get_filename()}")
                    save_attachment_in_memory(part, mail_item)
                elif content_type == 'text/plain':
                    mail_item.Body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                elif content_type == 'text/html':
                    mail_item.HTMLBody = part.get_payload(decode=True).decode('utf-8', errors='ignore')
        else:
            mail_item.Body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

        mail_item.Save()
        mail_item.Move(inbox_folder)
        logging.info(f"Email processed successfully.")
    except Exception as e:
        logging.error(f"Error processing email: {e}")
        logging.error(traceback.format_exc())  # Log stack trace for debugging

def batch_process_emails(emails, inbox_folder, output_folder, batch_size=500):
    """Processes emails in batches to avoid memory overload and improve performance."""
    print("Processing emails in batches...")
    for i in tqdm(range(0, len(emails), batch_size), desc="Processing batches"):
        batch = emails[i:i + batch_size]
        with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as executor:
            futures = [executor.submit(process_email_with_retry, email, inbox_folder, output_folder) for email in batch]
            concurrent.futures.wait(futures)
        
        save_checkpoint()  # Save progress after each batch
        logging.info(f"Processed batch {i // batch_size + 1}/{len(emails) // batch_size + 1}")
        gc.collect()  # Clean up memory

def ensure_directory_exists(pst_file):
    """Ensure the directory for the PST file exists. If not, create it."""
    pst_dir = os.path.dirname(pst_file)
    if pst_dir and not os.path.exists(pst_dir):  # Check if directory is specified and doesn't exist
        try:
            os.makedirs(pst_dir)  # Create the directory
            print(f"Created directory for PST: {pst_dir}")
        except OSError as e:
            logging.error(f"Error creating directory {pst_dir}: {e}")
            raise

def backup_existing_pst(pst_file):
    """Backup existing PST file to prevent overwriting."""
    if os.path.exists(pst_file):
        backup_file = pst_file + ".backup"
        shutil.copy(pst_file, backup_file)
        print(f"Backed up existing PST file to: {backup_file}")

@retry(wait_exponential_multiplier=1000, wait_exponential_max=10000, stop_max_attempt_number=5)
def import_emails_to_outlook(emails, pst_file, output_folder):
    """Imports the processed emails into Outlook PST file."""
    try:
        outlook = ensure_outlook_running()

        ensure_directory_exists(pst_file)
        backup_existing_pst(pst_file)
        
        namespace = outlook.GetNamespace("MAPI")
        print(f"Creating PST file: {pst_file}")
        logging.info(f"Creating PST file: {pst_file}")
        namespace.AddStoreEx(pst_file, 3)  # Create new PST file
        pst_folder = namespace.Folders.GetLast()

        inbox_folder = pst_folder.Folders.Add("Inbox") if not get_folder_by_name(pst_folder, "Inbox") else pst_folder.Folders["Inbox"]

        batch_process_emails(emails, inbox_folder, output_folder)

    except Exception as e:
        logging.error(f"Error creating or importing PST file: {e}")
        logging.error(traceback.format_exc())  # Log stack trace for debugging
        raise

def get_folder_by_name(parent_folder, folder_name):
    """Get a folder by name from the parent folder, or return None if not found."""
    for folder in parent_folder.Folders:
        if folder.Name == folder_name:
            return folder
    return None

def confirm_file_overwrite(pst_file):
    """Interactive confirmation for overwriting existing PST file."""
    if os.path.exists(pst_file):
        confirmation = input(f"The PST file '{pst_file}' already exists. Do you want to overwrite it? (y/n): ")
        if confirmation.lower() != 'y':
            print("Operation cancelled.")
            sys.exit(0)

def generate_summary_report(total_emails, failed_emails, start_time):
    """Generate a summary report at the end of the process."""
    elapsed_time = time.time() - start_time
    print(f"\nSummary Report:")
    print(f"Total Emails Processed: {total_emails}")
    print(f"Failed Emails: {failed_emails}")
    print(f"Time Taken: {elapsed_time:.2f} seconds")

if __name__ == "__main__":
    args = setup_logging_and_parse_args()

    # Initialize logging and load processed emails checkpoint
    processed_emails = initialize_logging_and_checkpoint(args.log_level)

    # If no MBOX file is provided, auto-detect .mbox files in the current directory
    if not args.mbox_file:
        mbox_files = find_mbox_files()
        if not mbox_files:
            print("No .mbox files found in the current directory.")
            sys.exit(1)
        else:
            print(f"Found .mbox files: {mbox_files}")
            args.mbox_file = mbox_files[0]  # Process the first .mbox file found

    # Check if the MBOX file exists
    if not os.path.exists(args.mbox_file):
        print(f"Error: File '{args.mbox_file}' does not exist.")
        sys.exit(1)

    # Confirm file overwrite before proceeding
    confirm_file_overwrite(args.pst_file)

    # Determine the directory of the MBOX file to create a PST file
    pst_file = os.path.abspath(args.pst_file)

    # Start time tracking for the summary report
    start_time = time.time()

    # Extract emails from the MBOX file using streaming
    emails = list(extract_emails_from_mbox_stream(args.mbox_file))

    # Import the emails into a new PST file in Outlook
    import_emails_to_outlook(emails, pst_file, args.output_folder)

    # Generate a summary report at the end
    generate_summary_report(len(emails), len(processed_emails), start_time)

    print(f"Conversion completed. PST saved at {pst_file}")
