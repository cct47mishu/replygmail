import os
import sys
import base64
import time
from datetime import date
import random
import pdfkit
from io import BytesIO
from email.mime.text import MIMEText
from email.message import EmailMessage
from colorama import Fore, init
from threading import Lock
from concurrent.futures import ThreadPoolExecutor
import email
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os
import imaplib
import smtplib
import ssl
import httplib2
init(strip=not sys.stdout.isatty())
init(autoreset=True)
ssl_context = ssl.create_default_context()

# Configure httplib2 with SSL context
http = httplib2.Http()
http.disable_ssl_certificate_validation = False  # Ensure SSL validation is enabled

Sent = 0
Error = 0
MAX_THREADS = 1  # Set max threads for concurrent email sending
lock = Lock()

if sys.platform == "win32":
    path_to_wkhtmltopdf = r'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'
else:
    path_to_wkhtmltopdf = '/usr/local/bin/wkhtmltopdf'

pdfkit_config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)

def sanitize_header(header_value):
    """ Remove any newline or carriage return characters from a header value """
    return header_value.replace("\n", "").replace("\r", "")

def load_variables():
    with open("tools/NAME_Var.txt") as f:
        names = f.read().splitlines()
    with open("tools/INVOCIE_Var.txt") as f:
        invoices = f.read().splitlines()
    with open("tools/PRODUCT_Var.txt") as f:
        products = f.read().splitlines()
    with open("tools/NUMBER_Var.txt") as f:
        numbers = f.read().splitlines()
    with open("tools/TEXT_Var.txt") as f:
        texts = f.read().splitlines()
    with open("tools/Signetures_Var.txt") as f:
        signatures = f.read().splitlines()
    with open("tools/body.txt", encoding="utf-8") as f:
        body = f.read()
    with open("tools/body2.txt", encoding="utf-8") as f:
        body2 = f.read()
    return names, invoices, products, numbers, texts, signatures, body, body2

names, invoices, products, numbers, texts, signatures, body, body2 = load_variables()

def generate_pdf_from_html(html_content, output):
    pdf_data = pdfkit.from_string(html_content, False, configuration=pdfkit_config)
    return BytesIO(pdf_data), f"{output}.pdf"

auth_method = input("[+] Choose authentication method (1 for App Password, 2 for Gmail API) : ")
time_wait = int(input("[+] Enter Time To Wait : "))

if auth_method == '1':
    EMAIL_ADDRESS = input("[+] Enter your email address : ")
    APP_PASSWORD = input("[+] Enter your app password : ")

    def check_inbox_imap():
        global Sent, Error
        try:
            mail = imaplib.IMAP4_SSL('imap.gmail.com')
            mail.login(EMAIL_ADDRESS, APP_PASSWORD)
            mail.select('inbox')
            status, messages = mail.search(None, 'UNSEEN')
            email_ids = messages[0].split()

            if email_ids:
                with lock:
                    print(Fore.LIGHTCYAN_EX + f"{EMAIL_ADDRESS} >> Has {len(email_ids)} New Messages :)")
                with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
                    for email_id in email_ids:
                        executor.submit(process_email_imap, mail, email_id)
            else:
                with lock:
                    print(Fore.LIGHTYELLOW_EX + "No new messages.")
            mail.logout()
        except Exception as e:
            with lock:
                Error += 1
                print(Fore.LIGHTRED_EX + f"Error in check_inbox_imap: {e}")

    def process_email_imap(mail, email_id):
        global Sent, Error
        try:
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    email_from = sanitize_header(msg['from'])
                    email_subject = sanitize_header(msg['subject'])
                    message_id = sanitize_header(msg['Message-ID'])
                    reply_to = msg.get('Reply-To')
                    to_email = sanitize_header(reply_to if reply_to else email_from)
                    with lock:
                        print(Fore.LIGHTWHITE_EX + f"\tNew email from: {email_from}")
                        print(Fore.LIGHTWHITE_EX + f"\tSubject: {email_subject}")
                    send_reply_imap(to_email, email_subject, message_id)
                    mail.store(email_id, '+FLAGS', '\\Seen')
        except Exception as e:
            with lock:
                Error += 1
                print(Fore.LIGHTRED_EX + f"Error processing email: {e}")

    def send_reply_imap(to_email, original_subject, message_id):
        global Sent, Error
        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_ADDRESS, APP_PASSWORD)

                today = date.today()
                emailMsg = body.replace("#Signetures_Var#", random.choice(signatures))\
                               .replace("#NAME#", random.choice(names))\
                               .replace("#INVOCIE#", random.choice(invoices))\
                               .replace("#DATE#", str(today))\
                               .replace("#PRODUCT#", random.choice(products))\
                               .replace("#NUMBER#", random.choice(numbers))\
                               .replace("#TEXT#", random.choice(texts))

                emailMsgs = body2.replace("#Signetures_Var#", random.choice(signatures))\
                                 .replace("#NAME#", random.choice(names))\
                                 .replace("#INVOCIE#", random.choice(invoices))\
                                 .replace("#DATE#", str(today))\
                                 .replace("#PRODUCT#", random.choice(products))\
                                 .replace("#NUMBER#", random.choice(numbers))\
                                 .replace("#TEXT#", random.choice(texts))

                msg = EmailMessage()
                msg['From'] = EMAIL_ADDRESS
                msg['To'] = sanitize_header(to_email)
                msg['Subject'] = sanitize_header(f"Re: {original_subject}")
                msg['In-Reply-To'] = sanitize_header(message_id)
                msg['References'] = sanitize_header(message_id)
                msg.add_alternative(emailMsg, subtype='html')

                pdf_data, pdf_filename = generate_pdf_from_html(emailMsgs, random.randint(999999999999, 9999999999999))
                msg.add_attachment(pdf_data.getvalue(), maintype='application', subtype='pdf', filename=pdf_filename)

                smtp.send_message(msg)
                with lock:
                    print(Fore.LIGHTGREEN_EX + f"Reply sent to {to_email}")
                    Sent += 1
        except Exception as e:
            with lock:
                Error += 1
                print(Fore.LIGHTRED_EX + f"Failed to send reply to {to_email}: {e}")

    while True:
        check_inbox_imap()
        with lock:
            print(Fore.YELLOW + "Waiting for the next check...")
        time.sleep(time_wait)

elif auth_method == '2':
    def authenticate_gmail():
        SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
        creds = None
        if os.path.exists('tools/token.json'):
            creds = Credentials.from_authorized_user_file('tools/token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                from google_auth_oauthlib.flow import InstalledAppFlow
                flow = InstalledAppFlow.from_client_secrets_file('tools/credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('tools/token.json', 'w') as token:
                token.write(creds.to_json())

        service = build('gmail', 'v1', credentials=creds, cache_discovery=False)
        return service

    def check_inbox_api(service):
        global Sent, Error
        try:
            results = service.users().messages().list(userId='me', q='is:unread').execute()
            messages = results.get('messages', [])

            if not messages:
                with lock:
                    print(Fore.LIGHTYELLOW_EX + "No new messages.")
                return

            with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
                for message in messages:
                    executor.submit(process_email_api, service, message['id'])
        except Exception as e:
            with lock:
                Error += 1
                print(Fore.LIGHTRED_EX + f"Error in check_inbox_api: {e}")

    def process_email_api(service, message_id):
        global Sent, Error
        try:
            msg = service.users().messages().get(userId='me', id=message_id).execute()
            email_data = msg['payload']['headers']
            email_from = sanitize_header(next(header['value'] for header in email_data if header['name'] == 'From'))
            email_subject = sanitize_header(next(header['value'] for header in email_data if header['name'] == 'Subject'))
            reply_to = next((header['value'] for header in email_data if header['name'].lower() == 'reply-to'), None)
            if not reply_to:
                reply_to = next((header['value'] for header in email_data if header['name'].lower() == 'from'), None)

            print("Replying to:", reply_to)

            to_email = sanitize_header(reply_to if reply_to else email_from)

            with lock:
                print(Fore.LIGHTCYAN_EX + f"New email from: {to_email}, Subject: {email_subject}")
            send_reply_api(service, to_email, email_subject, message_id)

            service.users().messages().modify(userId='me', id=message_id, body={'removeLabelIds': ['UNREAD']}).execute()
        except Exception as e:
            with lock:
                Error += 1
                print(Fore.LIGHTRED_EX + f"Error processing email with Gmail API: {e}")

    def send_reply_api(service, to_email, original_subject, message_id):
        global Sent, Error
        try:
            today = date.today()
            emailMsg = body.replace("#Signetures_Var#", random.choice(signatures))\
                            .replace("#NAME#", random.choice(names))\
                            .replace("#INVOCIE#", random.choice(invoices))\
                            .replace("#DATE#", str(today))\
                            .replace("#PRODUCT#", random.choice(products))\
                            .replace("#NUMBER#", random.choice(numbers))\
                            .replace("#TEXT#", random.choice(texts))

            emailMsgs = body2.replace("#Signetures_Var#", random.choice(signatures))\
                                .replace("#NAME#", random.choice(names))\
                                .replace("#INVOCIE#", random.choice(invoices))\
                                .replace("#DATE#", str(today))\
                                .replace("#PRODUCT#", random.choice(products))\
                                .replace("#NUMBER#", random.choice(numbers))\
                                .replace("#TEXT#", random.choice(texts))

            msg = EmailMessage()
            msg['To'] = sanitize_header(to_email)
            msg['Subject'] = sanitize_header(f"Re: {original_subject}")
            msg['In-Reply-To'] = sanitize_header(message_id)
            msg['References'] = sanitize_header(message_id)
            msg.add_alternative(emailMsg, subtype='html')

            pdf_data, pdf_filename = generate_pdf_from_html(emailMsgs, random.randint(999999999999, 9999999999999))
            msg.add_attachment(pdf_data.getvalue(), maintype='application', subtype='pdf', filename=pdf_filename)

            raw_message = {'raw': base64.urlsafe_b64encode(msg.as_bytes()).decode()}
            service.users().messages().send(userId='me', body=raw_message).execute()

            with lock:
                print(Fore.LIGHTGREEN_EX + f"Reply sent to {to_email}")
                Sent += 1
        except Exception as e:
            with lock:
                Error += 1
                print(Fore.LIGHTRED_EX + f"Failed to send reply to {to_email}: {e}")

    service = authenticate_gmail()
    while True:
        check_inbox_api(service)
        with lock:
            print(Fore.YELLOW + "Waiting for the next check...")
        time.sleep(time_wait)

else:
    print("Invalid choice. Please run the script again and choose either 1 or 2.")
