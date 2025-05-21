import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ─────────────────────────────────────
# 1)  SEND A SINGLE .DOCX ATTACHMENT
# ─────────────────────────────────────
def send_invoice_email(sender_email, sender_password,
                       recipient_email, subject, body,
                       attachment_path):
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"]   = recipient_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    # Attach the .docx file
    with open(attachment_path, "rb") as f:
        part = MIMEApplication(
            f.read(),
            _subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            Name=os.path.basename(attachment_path)
        )
    part["Content-Disposition"] = (
        f'attachment; filename="{os.path.basename(attachment_path)}"'
    )
    msg.attach(part)

    # Send the email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, sender_password)
        server.send_message(msg)
        print(f"Sent invoice to {recipient_email}")

# ─────────────────────────────────────
# 2)  LOOP THROUGH RECORDS & SEND .DOCX
# ─────────────────────────────────────
def send_invoices_for_records(processed_records, output_dir,
                              sender_email, sender_password):
    for record in processed_records:
        recipient_email = record.get("email")
        invoice_number  = record.get("Invoice Number", "invoice")
        docx_path = os.path.join(output_dir, f"{invoice_number}.docx")

        if os.path.exists(docx_path) and recipient_email:
            subject = f"Your Invoice {invoice_number}"
            body = (
                f"Dear {record.get('Beneficiary Name', '')},\n\n"
                "Please find attached your invoice.\n\n"
                "Best regards,\nEZ Works"
            )
            send_invoice_email(sender_email, sender_password,
                               recipient_email, subject, body,
                               docx_path)          # ⬅ attach DOCX
        else:
            print(f"Invoice file or email missing for {invoice_number}")
