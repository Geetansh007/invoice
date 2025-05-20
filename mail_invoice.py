import smtplib
import os
import subprocess
from email.message import EmailMessage
from email.utils import make_msgid
from email.mime.base import MIMEBase
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def docx_to_pdf(docx_path, output_dir):
    # Use libreoffice to convert docx to pdf
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path
    ], check=True)
    pdf_path = os.path.join(output_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    return pdf_path

def send_invoice_email(sender_email, sender_password, recipient_email, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    # Attach the invoice file (PDF)
    with open(attachment_path, 'rb') as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
    msg.attach(part)

    # Send the email via SMTP
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(sender_email, sender_password)
        server.send_message(msg)
        print(f"Sent invoice to {recipient_email}")


def send_invoices_for_records(processed_records, output_dir, sender_email, sender_password):
    for record in processed_records:
        recipient_email = record.get('email')
        invoice_number = record.get('Invoice Number', 'invoice')
        docx_path = os.path.join(output_dir, f"{invoice_number}.docx")
        if os.path.exists(docx_path) and recipient_email:
            pdf_path = docx_to_pdf(docx_path, output_dir)
            subject = f"Your Invoice {invoice_number}"
            body = f"Dear {record.get('Beneficiary Name', '')},\n\nPlease find attached your invoice.\n\nBest regards,\nEZ Works"
            send_invoice_email(sender_email, sender_password, recipient_email, subject, body, pdf_path)
        else:
            print(f"Invoice or email missing for {invoice_number}") 