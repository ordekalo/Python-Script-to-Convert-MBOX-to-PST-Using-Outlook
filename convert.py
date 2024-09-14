import sys
import mailbox
import os
import email
import win32com.client
import pythoncom
from tqdm import tqdm
import logging
import concurrent.futures
import gc
import time
import hashlib
import argparse
import traceback
from retrying import retry

# Initialize logging
logging.basicConfig(filename='import_log.txt', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

processed_emails = set()

def parse_args():
    """
    Parse command-line arguments.
    """
    parser = argparse.ArgumentParser(description="Convert MBOX to PST using Outlook")
    parser.add_argument("--mbox_file", help="Path to the MBOX file. If not provided, the script will auto-detect all '.mbox' files in the current directory.")
    parser.add_argument("output_folder", help="Folder to save the attachments")
    parser.add_argument("--pst_file", help="Path to the PST file", default="emails.pst")
    return parser.parse_args()

def find_mbox_files():
    """
    Find all .mbox files in the current directory.
    """
    mbox_files = [f for f in os.listdir('.') if f.endswith('.mbox')]
    return mbox_files

def hash_email(raw_email):
    """
    Create a unique hash of the email content to track processed emails.
    """
    return hashlib.sha256(raw_email.encode('utf-8')).hexdigest()

def extract_emails_from_mbox(mbox_file):
    """
    Extract emails from an MBOX file and return them as a list of raw email strings.
    """
    try:
        print(f"Opening MBOX file: {mbox_file}")
        mbox = mailbox.mbox(mbox_file)
        emails = [message.as_string() for message in mbox]
        print(f"Extracted {len(emails)} emails from the MBOX file.")
        return emails
    except Exception as e:
        logging.error(f"Error extracting emails from MBOX: {e}")
        raise

def save_attachment(part, mail_item):
    """
    Save the attachment from the email part to the Outlook mail item.
    """
    filename = part.get_filename()
    if filename:
        payload = part.get_payload(decode=True)
        if payload:
            # Save attachment to the mail item
            with open(filename, 'wb') as f:
                f.write(payload)
            mail_item.Attachments.Add(os.path.abspath(filename))
            os.remove(filename)  # Clean up after attaching

@retry(wait_exponential_multiplier=1000, wait_exponential_max=10000, stop_max_attempt_number=5)
def ensure_outlook_running():
    """
    Ensure that Outlook is running and accessible.
    Retry if Outlook is not available with exponential backoff.
    """
    try:
        pythoncom.CoInitialize()  # Initialize COM library
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as e:
        logging.error(f"Failed to connect to Outlook: {e}")
        raise

def release_com_object(obj):
    """
    Properly release a COM object to avoid memory leaks.
    """
    if obj:
        obj = None

def process_email_with_retry(raw_email, inbox_folder, output_folder, retries=3):
    """
    Attempts to process an email and import it into Outlook.
    Retries up to 'retries' times if an error occurs.
    """
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
                    save_attachment(part, mail_item)
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
    """
    Processes emails in batches to avoid memory overload and improve performance.
    """
    for i in range(0, len(emails), batch_size):
        batch = emails[i:i + batch_size]
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(process_email_with_retry, email, inbox_folder, output_folder) for email in batch]
            concurrent.futures.wait(futures)
        
        logging.info(f"Processed batch {i // batch_size + 1}/{len(emails) // batch_size + 1}")
        gc.collect()  # Clean up memory

def ensure_directory_exists(pst_file):
    """
    Ensure the directory for the PST file exists. If not, create it.
    """
    pst_dir = os.path.dirname(pst_file)
    if pst_dir and not os.path.exists(pst_dir):  # Check if directory is specified and doesn't exist
        try:
            os.makedirs(pst_dir)  # Create the directory
            print(f"Created directory for PST: {pst_dir}")
        except OSError as e:
            logging.error(f"Error creating directory {pst_dir}: {e}")
            raise

@retry(wait_exponential_multiplier=1000, wait_exponential_max=10000, stop_max_attempt_number=5)
def import_emails_to_outlook(emails, pst_file, output_folder):
    try:
        # Ensure Outlook is running
        outlook = ensure_outlook_running()

        # Ensure the directory for the PST file exists
        ensure_directory_exists(pst_file)
        
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
    """
    Get a folder by name from the parent folder, or return None if not found.
    """
    for folder in parent_folder.Folders:
        if folder.Name == folder_name:
            return folder
    return None

if __name__ == "__main__":
    args = parse_args()

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

    # Determine the directory of the MBOX file to create a PST file
    pst_file = os.path.abspath(args.pst_file)
    
    # Extract emails from the MBOX file
    emails = extract_emails_from_mbox(args.mbox_file)

    # Import the emails into a new PST file in Outlook
    import_emails_to_outlook(emails, pst_file, args.output_folder)

    print(f"Conversion completed. PST saved at {pst_file}")
