import msal
import requests
import json
import base64
import os
from datetime import datetime

# --- CONFIGURATION ---
CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "2362dd87-1d36-4d76-b592-5daf74ac1a1d")
TENANT_ID = os.getenv("GRAPH_TENANT_ID", "743a5d9f-1123-4f3f-8fcf-5766b8ad8bf9")
CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "")

# The mailbox we are accessing (Reading from AND Sending as)
TARGET_MAILBOX = 'BRProdUKCOP03@gammatelecom.com' 

# Who is receiving the test email?
RECIPIENT_EMAIL = 'david.murphy+emailtest@gamma.co.uk' 

# Graph API Scopes
SCOPE = ['https://graph.microsoft.com/.default'] 

def get_service_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    
    result = app.acquire_token_for_client(scopes=SCOPE)

    if "access_token" in result:
        print("[OK] Authentication successful! Service Token acquired.")
        return result['access_token']
    else:
        print(f"[ERROR] Authentication failed: {result.get('error_description')}")
        return None

def get_latest_email(access_token):
    print(f"\n--- Attempting to READ the last email from {TARGET_MAILBOX} ---")
    
    # URL to get messages from Inbox, sorted by newest first, take top 1
    read_endpoint = f'https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/mailFolders/inbox/messages'
    params = {
        '$top': 1,
        '$orderby': 'receivedDateTime desc',
        '$select': 'subject,from,receivedDateTime'
    }
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.get(read_endpoint, headers=headers, params=params)

    if response.status_code == 200:
        data = response.json()
        messages = data.get('value', [])
        
        if messages:
            msg = messages[0]
            print("[OK] SUCCESS: Latest email found!")
            print(f"   Subject: {msg.get('subject')}")
            print(f"   From:    {msg.get('from', {}).get('emailAddress', {}).get('address')}")
            print(f"   Date:    {msg.get('receivedDateTime')}")
        else:
            print("[INFO] Inbox is empty.")
    elif response.status_code == 403:
        print("[ERROR] FAILED TO READ: Access Denied (403).")
        print("   Reason: The App likely lacks 'Mail.Read' permissions in Azure or RBAC.")
    else:
        print(f"[ERROR] FAILED TO READ: Status {response.status_code}")
        print(response.text)

def send_test_email(access_token):
    print(f"\n--- Attempting to SEND email to {RECIPIENT_EMAIL} ---")
    
    send_endpoint = f'https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/sendMail'
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    email_data = {
        "message": {
            "subject": "Automation Test: Read & Send Verification",
            "body": {
                "contentType": "Text",
                "content": f"This is a test. The script attempted to read the latest email from {TARGET_MAILBOX} and then sent this message."
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": RECIPIENT_EMAIL
                    }
                }
            ]
        }
    }

    response = requests.post(send_endpoint, headers=headers, json=email_data)

    if response.status_code == 202:
        print("[OK] SUCCESS: Test email sent successfully!")
    else:
        print(f"[ERROR] FAILED TO SEND: Status {response.status_code}")
        print(response.text)

def send_order_report_email(report_file_path, recipient_email='david.murphy+psmsoutput@gamma.co.uk', 
                           input_filename='', total_orders=0, successful=0, failed=0, errors=0,
                           report_type='Order Status', subject_prefix='SoGEA Order Status Report'):
    """
    Send an order report via email with attachment
    
    Args:
        report_file_path: Full path to the Excel report file
        recipient_email: Recipient email(s) as string (comma/semicolon-separated) or list
        input_filename: Name of the input file that was processed
        total_orders: Total number of orders checked/processed
        successful: Number of successful operations
        failed: Number of failed operations
        errors: Number of errors
        report_type: Type of report (default: 'Order Status')
        subject_prefix: Prefix for email subject (default: 'SoGEA Order Status Report')
    
    Returns:
        bool: True if email sent successfully, False otherwise
    """
    print(f"\n--- Sending order report to {recipient_email} ---")

    # Build recipient list (supports comma/semicolon-separated string or list)
    if isinstance(recipient_email, str):
        raw_addresses = recipient_email.replace(';', ',').split(',')
        recipient_addresses = [addr.strip() for addr in raw_addresses if addr and addr.strip()]
    elif isinstance(recipient_email, (list, tuple, set)):
        recipient_addresses = [str(addr).strip() for addr in recipient_email if str(addr).strip()]
    else:
        recipient_addresses = [str(recipient_email).strip()] if str(recipient_email).strip() else []

    if not recipient_addresses:
        print("[ERROR] No valid recipient email addresses provided")
        return False
    
    # Get authentication token
    token = get_service_token()
    if not token:
        print("[ERROR] Failed to authenticate")
        return False
    
    # Check if file path is provided and file exists
    if report_file_path is None:
        print(f"[ERROR] No report file path provided")
        return False
    
    if not os.path.exists(report_file_path):
        print(f"[ERROR] Report file not found: {report_file_path}")
        return False
    
    # Read file and encode to base64
    try:
        with open(report_file_path, 'rb') as f:
            file_content = f.read()
            file_base64 = base64.b64encode(file_content).decode('utf-8')
    except Exception as e:
        print(f"[ERROR] Failed to read file: {e}")
        return False
    
    filename = os.path.basename(report_file_path)
    
    # Build email body
    email_body = f"""{subject_prefix}

Attached is the automated {report_type.lower()} report.

Input File: {input_filename}
Total Orders: {total_orders}
Successful: {successful}
Failed: {failed}
Errors: {errors}

Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

This is an automated message from the PSTN Switch-off automation system.
"""
    
    send_endpoint = f'https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/sendMail'
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    email_data = {
        "message": {
            "subject": f"{subject_prefix} - {input_filename or filename}",
            "body": {
                "contentType": "Text",
                "content": email_body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": address
                    }
                }
                for address in recipient_addresses
            ],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename,
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "contentBytes": file_base64
                }
            ]
        }
    }

    response = requests.post(send_endpoint, headers=headers, json=email_data)

    if response.status_code == 202:
        print(f"[OK] SUCCESS: Order report emailed to {', '.join(recipient_addresses)}")
        return True
    else:
        print(f"[ERROR] FAILED TO SEND: Status {response.status_code}")
        print(response.text)
        return False

if __name__ == "__main__":
    try:
        if not CLIENT_SECRET or 'PASTE' in CLIENT_SECRET:
            print("⚠️ STOP: Please insert your Client Secret in the script configuration.")
        else:
            token = get_service_token()
            if token:
                # 1. Try to Read
                get_latest_email(token)
                # 2. Try to Send
                send_test_email(token)
    except Exception as e:
        print(f"An error occurred: {e}")