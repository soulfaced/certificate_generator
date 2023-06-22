import io
import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
import base64
# import json
# import requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import mimetypes
from google.auth.transport.requests import Request
from PyPDF2 import PdfWriter, PdfReader

certemplate = "template.pdf"
excelfile = "data.xlsx"
varname = "Name"
horz = 310
vert = 275
varfont = "Inter.ttf"
fontsize = 20
fontcolor = "#ffffff"
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "padwekarsanchit@gmail.com"  # Your email address
credentials_file = "cred.json"
token_file = "token.json"

# Create the certificate directory
os.makedirs("certificates", exist_ok=True)

# Register the necessary font
pdfmetrics.registerFont(TTFont('myFont', varfont))

# Read the data from the Excel file
data = pd.read_excel(excelfile)
name_list = data[varname].tolist()
names = [str(x).strip().upper() for x in name_list]

# Create PDF certificates
for i in names:
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # the registered font is used, provide font size and color
    can.setFont("myFont", fontsize)
    can.setFillColor(HexColor(fontcolor))  # Set font color

    # provide the text location in pixels
    can.drawString(horz, vert, i)

    can.save()
    packet.seek(0)
    new_pdf = PdfReader(packet)

    # provide the certificate template
    existing_pdf = PdfReader(open(certemplate, "rb"))

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    destination = "certificates" + os.sep + i + ".pdf"
    outputStream = open(destination, "wb")
    output.write(outputStream)
    outputStream.close()
    print("created " + i + ".pdf")

# Authenticate Gmail
def authenticate_gmail():
    creds = None
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, ["https://www.googleapis.com/auth/gmail.modify"])
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, ["https://www.googleapis.com/auth/gmail.modify"])
            creds = flow.run_local_server(port=0)
        with open(token_file, "w") as token:
            token.write(creds.to_json())
    return creds

# Send email with the certificate attached using Gmail API
def send_email_with_attachment(service, sender_email, recipient_email, subject, body, attachment_path):
    message = MIMEMultipart()
    message["to"] = recipient_email
    message["from"] = sender_email
    message["subject"] = subject

    message.attach(MIMEText(body, "plain"))

    content_type, encoding = mimetypes.guess_type(attachment_path)
    if content_type is None or encoding is not None:
        content_type = "application/octet-stream"

    main_type, sub_type = content_type.split("/", 1)
    with open(attachment_path, "rb") as attachment:
        payload = MIMEBase(main_type, sub_type)
        payload.set_payload(attachment.read())
        encoders.encode_base64(payload)

    payload.add_header("Content-Disposition", "attachment", filename=os.path.basename(attachment_path))
    message.attach(payload)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    try:
        service.users().messages().send(userId="me", body={"raw": raw_message}).execute()
        print(f"Sent email to: {recipient_email}")
    except Exception as e:
        print(f"An error occurred while sending the email to {recipient_email}: {str(e)}")

# Authenticate Gmail API
credentials = authenticate_gmail()
service = build("gmail", "v1", credentials=credentials)

email_list = data['Email'].tolist()
emails = [str(x).strip() for x in email_list]

# Send emails with attached certificates
for name, email in zip(names, emails):
    subject = "Certificate of Achievement"
    body = f"Dear {name},\n\nPlease find attached your certificate of achievement.\n\nBest regards,\nYour Name"
    attachment_path = os.path.join("certificates", f"{name}.pdf")
    send_email_with_attachment(service, sender_email, email, subject, body, attachment_path)
