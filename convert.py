import sys
import mailbox
import os
import email
import win32com.client

def extract_emails_from_mbox(mbox_file):
    """
    Extract emails from an MBOX file and return them as a list of raw email strings.
    """
    mbox = mailbox.mbox(mbox_file)
    emails = []
    for message in mbox:
        emails.append(message.as_string())
    return emails

def get_folder_by_name(parent_folder, folder_name):
    """
    Get a folder by name from the parent folder, or return None if not found.
    """
    for folder in parent_folder.Folders:
        if folder.Name == folder_name:
            return folder
    return None

def import_emails_to_outlook(emails, pst_file):
    """
    Import a list of emails into a new PST file in Outlook.
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Add a new PST file to Outlook
    namespace.AddStoreEx(pst_file, 3)  # 3 = Unicode PST format
    pst_folder = namespace.Folders.GetLast()

    # Create "Inbox" folder if it doesn't exist
    inbox_folder = get_folder_by_name(pst_folder, "Inbox")
    if inbox_folder is None:
        inbox_folder = pst_folder.Folders.Add("Inbox")

    for raw_email in emails:
        msg = email.message_from_string(raw_email)

        # Create a new mail item in Outlook
        mail_item = outlook.CreateItem(0)  # 0 = olMailItem
        mail_item.Subject = msg['subject'] or "(No Subject)"
        mail_item.BodyFormat = 1  # 1 = plain text format

        # Handle the email body (text/plain or text/html)
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                payload = part.get_payload(decode=True)
                if content_type == 'text/plain' and payload:
                    mail_item.Body = payload.decode('utf-8', errors='ignore')
                elif content_type == 'text/html' and payload:
                    mail_item.HTMLBody = payload.decode('utf-8', errors='ignore')
        else:
            content_type = msg.get_content_type()
            payload = msg.get_payload(decode=True)
            if content_type == 'text/plain' and payload:
                mail_item.Body = payload.decode('utf-8', errors='ignore')
            elif content_type == 'text/html' and payload:
                mail_item.HTMLBody = payload.decode('utf-8', errors='ignore')

        # Save the mail item in the "Inbox" folder
        mail_item.Save()
        mail_item.Move(inbox_folder)

if __name__ == "__main__":
    # Ensure the correct number of arguments are provided
    if len(sys.argv) != 2:
        print("Usage: python convert.py <mbox_file>")
        sys.exit(1)

    # Get the MBOX file path from the command-line argument
    mbox_file = sys.argv[1]

    # Check if the MBOX file exists
    if not os.path.exists(mbox_file):
        print(f"Error: File '{mbox_file}' does not exist.")
        sys.exit(1)

    # Determine the directory of the MBOX file to create a PST file in the same directory
    pst_file = os.path.join(os.path.dirname(mbox_file), 'emails.pst')

    # Extract emails from the MBOX file
    emails = extract_emails_from_mbox(mbox_file)

    # Import the emails into a new PST file in Outlook
    import_emails_to_outlook(emails, pst_file)

    print(f"Conversion completed. PST saved at {pst_file}")
