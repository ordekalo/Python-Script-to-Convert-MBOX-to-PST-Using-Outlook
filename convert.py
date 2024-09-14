import sys
import mailbox
import os
import email
import win32com.client
from tqdm import tqdm
import logging
import concurrent.futures
import gc
import time
import argparse
import hashlib

# Initialize logging
logging.basicConfig(filename='import_log.txt', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

processed_emails = set()

def parse_args():
    parser = argparse.ArgumentParser(description="Convert MBOX to PST using Outlook")
    parser.add_argument("mbox_file", help="Path to the MBOX file")
    parser.add_argument("output_folder", help="Folder to save the attachments")
    parser.add_argument("--pst_file", help="Path to the PST file", default="emails.pst")
    return parser.parse_args()

def extract_emails_from_mbox(mbox_file):
    try:
        print(f"Opening MBOX file: {mbox_file}")
        mbox = mailbox.mbox(mbox_file)
        emails = [message.as_string() for message in mbox]
        print(f"Extracted {len(emails)} emails from the MBOX file.")
        return emails
    except Exception as e:
        logging.error(f"Error extracting emails: {e}")
        raise

def save_attachment_directly(part, mail_item):
    """
    Saves the attachment directly to the Outlook mail item without writing to disk.
    """
    filename = part.get_filename()
    if filename:
        payload = part.get_payload(decode=True)
        if payload:
            # Save the attachment directly to the Outlook email
            mail_item.Attachments.AddBytes(filename, payload)

def hash_email(raw_email):
    """
    Create a unique hash of the email content to track processed emails.
    """
    return hashlib.sha256(raw_email.encode('utf-8')).hexdigest()

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
        msg = email.message_from_string(raw_email, policy=email.policy.default)
        mail_item = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
        mail_item.Subject = msg['subject'] or "(No Subject)"
        mail_item.To = msg.get('to', '')
        mail_item.CC = msg.get('cc', '')
        mail_item.BCC = msg.get('bcc', '')
        mail_item.SentOn = msg.get('date', '')
        mail_item.SenderEmailAddress = msg.get('from', '')

        if msg.is_multipart():
            for part in msg.walk():
                if part.get_filename():
                    logging.info(f"Saving attachment: {part.get_filename()}")
                    save_attachment_directly(part, mail_item)
                elif part.get_content_type() == 'text/plain':
                    mail_item.Body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                elif part.get_content_type() == 'text/html':
                    mail_item.HTMLBody = part.get_payload(decode=True).decode('utf-8', errors='ignore')
        else:
            mail_item.Body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

        mail_item.Save()
        mail_item.Move(inbox_folder)
        logging.info(f"Email processed successfully.")
    except Exception as e:
        logging.error(f"Error processing email: {e}")

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

def import_emails_to_outlook(emails, pst_file, output_folder):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Ensure the directory for the PST file exists
        if not os.path.exists(os.path.dirname(pst_file)):
            os.makedirs(os.path.dirname(pst_file))
        
        print(f"Creating PST file: {pst_file}")
        logging.info(f"Creating PST file: {pst_file}")
        namespace.AddStoreEx(pst_file, 3)  # Create new PST file
        pst_folder = namespace.Folders.GetLast()

        inbox_folder = pst_folder.Folders.Add("Inbox") if not get_folder_by_name(pst_folder, "Inbox") else pst_folder.Folders["Inbox"]

        batch_process_emails(emails, inbox_folder, output_folder)
    except Exception as e:
        logging.error(f"Error creating or importing PST file: {e}")
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
    emails = extract_emails_from_mbox(args.mbox_file)
    import_emails_to_outlook(emails, args.pst_file, args.output_folder)
