import tkinter as tk
from tkinter import messagebox
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os
import time
import base64
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from dotenv import load_dotenv

load_dotenv()

COORDINATOR_DATA = {
    "rahul": {"fullName": "Rahul Kumar", "phone": "+91-7993185567", "email": "2024pgcsca022@nitjsr.ac.in"},
    "priya": {"fullName": "Priya Brahma", "phone": "+91-9678374608", "email": "2024pgcsca001@nitjsr.ac.in"},
    "rohit": {"fullName": "Rohit Gangwar", "phone": "+91-8979304759", "email": "2024pgcsca019@nitjsr.ac.in"},
    "siya": {"fullName": "Siya Agarwal", "phone": "+91-9310120759", "email": "2024pgcsca058@nitjsr.ac.in"},
    "taher": {"fullName": "Taher Mallik", "phone": "+91-9152760580", "email": "2024pgcsca007@nitjsr.ac.in"}
}

SCOPES = ['https://www.googleapis.com/auth/gmail.modify', 'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def authenticate():
    """Authenticates with Google and returns authorized service objects for Gmail and Sheets."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    gmail_service = build('gmail', 'v1', credentials=creds) 

    gspread_client = gspread.authorize(creds) 
    return gmail_service, gspread_client

def get_or_create_label_id(service, parent_label_name, sub_label_name):
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    parent_label_id, sub_label_id = None, None
    full_sub_label_name = f"{parent_label_name}/{sub_label_name}"

    for label in labels:
        if label['name'] == parent_label_name: parent_label_id = label['id']
        if label['name'] == full_sub_label_name: sub_label_id = label['id']

    if sub_label_id: return sub_label_id

    if not parent_label_id:
        label_body = {'name': parent_label_name, 'labelListVisibility': 'labelShow', 'messageListVisibility': 'show'}
        parent_label = service.users().labels().create(userId='me', body=label_body).execute()
        parent_label_id = parent_label['id']
    
    label_body = {'name': full_sub_label_name, 'labelListVisibility': 'labelShow', 'messageListVisibility': 'show'}
    sub_label = service.users().labels().create(userId='me', body=label_body).execute()
    return sub_label['id']


# Backend
def process_and_send_emails(sender_email, sheet_url, template_path, attachment_paths):
    try:
        gmail_service, gspread_client = authenticate() 
        
        sheet = gspread_client.open_by_url(sheet_url).sheet1
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        headers = sheet.row_values(1)
        sent_date_col = headers.index('Email Sent date') + 1

    except FileNotFoundError:
        messagebox.showerror("Auth Error", "credentials.json not found. Make sure it is the 'Desktop app' type.")
        return
    except Exception as e:
        messagebox.showerror("Authentication/Sheet Error", f"An error occurred: {e}")
        return

    
    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()

    emails_sent_count = 0
    for index, row in df.iterrows():
        if str(row['Email Sent date']).strip() == '':
            sender_key = str(row['Sent By']).strip().lower()
            coordinator_details = COORDINATOR_DATA.get(sender_key)
            if not coordinator_details:
                print(f"Warning: Coordinator '{row['Sent By']}' in row {index + 2} not found. Skipping.")
                continue

            try:
                label_id_to_apply = get_or_create_label_id(gmail_service, "MCA 2K27 Batch", coordinator_details['fullName'].split()[0])
            except Exception as e:
                messagebox.showerror("Label Error", f"Could not create or find Gmail label: {e}")
                return

            msg = MIMEMultipart()
            msg['From'] = f"{coordinator_details['fullName']} <{sender_email}>"
            msg['To'] = str(row['Recipient'])
            msg['Cc'] = str(row['Recipient(CC)']) if 'Recipient(CC)' in row and pd.notna(row['Recipient(CC)']) else ""
            msg['Subject'] = "Invitation: Campus Recruitment & Internships at NIT Jamshedpur"
            
            personalized_html = html_template.replace("{{HR name}}", str(row['HR name'])).replace("{{Company Name}}", str(row['Company Name'])).replace("{{Sent By Full Name}}", coordinator_details['fullName']).replace("{{Sent By Phone}}", coordinator_details['phone']).replace("{{Sent By Email}}", coordinator_details['email'])
            msg.attach(MIMEText(personalized_html, 'html'))
            
            for path in attachment_paths:
                part = MIMEBase('application', 'octet-stream')
                with open(path, 'rb') as file: part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(path)}"')
                msg.attach(part)

            try:
                encoded_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
                create_message_request = {'raw': encoded_message, 'labelIds': [label_id_to_apply]}
                sent_message = gmail_service.users().messages().send(userId='me', body=create_message_request).execute()
                print(f"Email sent to {row['Recipient']} with label. Message ID: {sent_message['id']}")
                
                sheet.update_cell(index + 2, sent_date_col, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                emails_sent_count += 1
                time.sleep(3)

            except Exception as e:
                messagebox.showwarning("Sending Error", f"Could not send email to {row['Recipient']}: {e}")

    messagebox.showinfo("Success", f"Mail merge complete! Sent {emails_sent_count} emails.")

class MailMergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gmail Mail Merge Automator")
        self.geometry("400x120")

        tk.Label(self, text="Configuration is loaded from project files.", font=("Helvetica", 10, "italic")).pack(pady=10)
        tk.Label(self, text="Click Start to begin.", font=("Helvetica", 10)).pack()
        tk.Button(self, text="START MAIL MERGE", bg="green", fg="white", font=("Helvetica", 12, "bold"), command=self.start_process).pack(pady=10, fill=tk.X, padx=20)
    
    def start_process(self):
        sender_email = os.getenv('SENDER_EMAIL')
        sheet_url = os.getenv('SHEET_URL')
        template_path = 'template.html'
        attachment_dir = 'attachments'

        if not os.path.isdir(attachment_dir):
            messagebox.showerror("Error", f"The '{attachment_dir}' directory was not found.")
            return
        attachment_paths = [os.path.join(attachment_dir, f) for f in os.listdir(attachment_dir) if os.path.isfile(os.path.join(attachment_dir, f))]
        
        if not all([sender_email, sheet_url]):
            messagebox.showwarning("Missing Information", "Please ensure your .env file is complete (email, URL).")
            return
        
        process_and_send_emails(sender_email, sheet_url, template_path, attachment_paths)


if __name__ == "__main__":
    app = MailMergeApp()
    app.mainloop()