import os
import imaplib
import email
import re
from email.header import decode_header
import pdfkit
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor

def html_to_pdf(html, pdf_filepath):
    # Replace this with the path to your wkhtmltopdf executable
    path_to_wkhtmltopdf = "C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"

    config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf.encode())

    # Parse the HTML using BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')

    # Remove problematic URLs
    for tag in soup.find_all(['img', 'a']):
        if tag.has_attr('href') or tag.has_attr('src'):
            try:
                url = tag['href'] if tag.has_attr('href') else tag['src']
                if not (url.startswith('http://') or url.startswith('https://')):
                    del tag['href']
            except KeyError:
                pass

    # Convert the modified HTML to a string
    modified_html = str(soup)

    pdfkit.from_string(modified_html, pdf_filepath, configuration=config)


def html_to_pdf(html, pdf_filepath):
    # Add a meta tag with the correct charset
    html = f'<meta charset="UTF-8">{html}'
    
    # Replace this with the path to your wkhtmltopdf executable
    path_to_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

    config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf.encode())
    pdfkit.from_string(html, pdf_filepath, configuration=config)

def sanitize_filename(filename):
    return re.sub(r'[\\/:*?"<>|]', '_', filename)


# Replace with your email and app password
email_address = "email"
app_password = "app password"

# Connect to the Gmail server using IMAP and SSL
imap = imaplib.IMAP4_SSL("imap.gmail.com")

# Log in using your email address and app password
imap.login(email_address, app_password)

# Select the mailbox you want to read from
imap.select("inbox")

# Search for starred (flagged) emails
status, response = imap.search(None, 'FLAGGED')

# Get the list of email IDs
email_ids = response[0].split()

# Create a directory to save email PDFs
os.makedirs("email_pdfs", exist_ok=True)

# Iterate through the email IDs and fetch each email
for email_id in email_ids:
    _, msg_data = imap.fetch(email_id, "(RFC822)")
    msg = email.message_from_bytes(msg_data[0][1])

    # Get the subject and decode it if necessary
    subject = decode_header(msg["Subject"])[0][0]
    if isinstance(subject, bytes):
        subject = subject.decode()

    print(f"Email ID: {email_id.decode()}")
    print(f"Subject: {subject}")

    # Initialize body variabledef process_email(email_id):
    _, msg_data = imap.fetch(email_id, "(RFC822)")
    msg = email.message_from_bytes(msg_data[0][1])

    # Get the subject and decode it if necessary
    subject = decode_header(msg["Subject"])[0][0]
    if isinstance(subject, bytes):
        subject = subject.decode()

    print(f"Email ID: {email_id.decode()}")
    print(f"Subject: {subject}")

    # Initialize body variable
    body = ""

    # Extract the email body
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                body = part.get_payload(decode=True).decode()
                break
    else:
        body = msg.get_payload(decode=True).decode()

    # Save the email content as a PDF
    sanitized_subject = sanitize_filename(subject)
    pdf_filename = f"{email_id.decode()}_{sanitized_subject}.pdf"
    pdf_filepath = os.path.join("email_pdfs", pdf_filename)
    html_to_pdf(body, pdf_filepath)

    print(f"Saved email as PDF: {pdf_filepath}")

# Close the mailbox and log out
imap.close()
imap.logout()
