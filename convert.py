import sys
import mailbox
import os
import email
import win32com.client
from tqdm import tqdm
import logging
import argparse

# Initialize logging
logging.basicConfig(filename='import_log.txt', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

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

def save_attachment(part, mail_item, output_folder):
    filename = part.get_filename()
    if filename:
        payload = part.get_payload(decode=True)
        if payload:
            attachment_path = os.path.join(output_folder, filename)
            with open(attachment_path, 'wb') as f:
                f.write(payload)
            mail_item.Attachments.Add(attachment_path)
            os.remove(attachment_path)  # Clean up

def process_email(raw_email, inbox_folder, output_folder):
    try:
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
                    save_attachment(part, mail_item, output_folder)
                elif part.get_content_type() == 'text/plain':
                    mail_item.Body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                elif part.get_content_type() == 'text/html':
                    mail_item.HTMLBody = part.get_payload(decode=True).decode('utf-8', errors='ignore')
        else:
            mail_item.Body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

        mail_item.Save()
        mail_item.Move(inbox_folder)
    except Exception as e:
        logging.error(f"Error processing email: {e}")

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

        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            executor.map(lambda email: process_email(email, inbox_folder, output_folder), emails)
    except Exception as e:
        logging.error(f"Error creating or importing PST file: {e}")
        raise

if __name__ == "__main__":
    args = parse_args()
    emails = extract_emails_from_mbox(args.mbox_file)
    import_emails_to_outlook(emails, args.pst_file, args.output_folder)
