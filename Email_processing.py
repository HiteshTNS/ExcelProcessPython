import io
import smtplib
from datetime import datetime
from typing import List, Dict, Any
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from openpyxl import Workbook


def generate_error_file(failed_invoice_rec_list: list) -> bytes:
    """
    Generates XLSX binary data for failed invoice records.
    Columns: Invoice #, Claim #, Status, Status Description
    """
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Failed Invoices"

    headers = ["Invoice #", "Claim #", "Status", "Status Description"]
    ws.append(headers)

    for record in failed_invoice_rec_list:
        row = [
            record.get("invoiceNumber", ""),
            record.get("claimNumber", ""),
            record.get("status", ""),
            record.get("statusDescription", "")
        ]
        ws.append(row)

    wb.save(output)
    return output.getvalue()


def save_xlsx_file(file_path: str, binary_data: bytes):
    """
    Saves binary XLSX data to a file, handling file lock errors gracefully.
    """
    try:
        with open(file_path, "wb") as file:
            file.write(binary_data)
        print(f"Excel file saved as '{file_path}'")
    except PermissionError:
        print(f"Cannot write to '{file_path}'. Please close the file if it's open and try again.")


def generate_invoice_email_body(file_detail: dict, attachment_name: str) -> str:
    """
    Generates HTML email body using file-related details.
    """
    uploaded_date = file_detail.get("uploadedDate", "").split("T")[0]
    return f"""
    <p>Please find the invoice file processing result as below:</p>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse;">
        <tr><td><strong>File Name</strong></td><td>{file_detail["fileName"]}</td></tr>
        <tr><td><strong>Uploaded By</strong></td><td>{file_detail["uploadedBy"]}</td></tr>
        <tr><td><strong>Uploaded Date</strong></td><td>{uploaded_date}</td></tr>
        <tr><td><strong>Total Records</strong></td><td>{file_detail["totalRecCount"]}</td></tr>
        <tr><td><strong>Error Records</strong></td><td>{file_detail["errorRecCount"]}</td></tr>
        <tr><td><strong>Success Records</strong></td><td>{file_detail["successRecCount"]}</td></tr>
        <tr><td><strong>Pending Records</strong></td><td>{file_detail["pendingRecCount"]}</td></tr>
        <tr><td><strong>Status</strong></td><td>{file_detail["status"]}</td></tr>
        <tr><td><strong>Status Description</strong></td><td>{file_detail["statusDescription"]}</td></tr>
    </table>
    <p>Please find the error invoice record details in attached spreadsheet: <strong>{attachment_name}</strong></p>
    <br>
    <p>Thanks,<br>VCIM SUPPORT</p>
    """


import os
import mimetypes
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib

def send_mail_notification(mail_config: dict):
    """
    Sends email with optional multiple attachments.
    Supports attachments via 'data' or file 'path'.

    Keys:
    - sender_email, sender_password
    - receiver_emails, cc_emails
    - subject, html_body
    - attachments: list of dicts:
        Either:
            - {'data': ..., 'filename': ..., 'mime_type': ...}
        Or:
            - {'path': 'path/to/file'}
    """
    sender = mail_config["sender_email"]
    password = mail_config["sender_password"]
    receivers = mail_config.get("receiver_emails", [])
    cc = mail_config.get("cc_emails", [])
    subject = mail_config["subject"]
    html_body = mail_config["html_body"]

    smtp_server = mail_config.get("smtp_server", "smtp.office365.com")
    smtp_port = mail_config.get("smtp_port", 587)

    message = MIMEMultipart()
    message["From"] = sender
    message["To"] = ", ".join(receivers)
    message["Cc"] = ", ".join(cc)
    message["Subject"] = subject
    message.attach(MIMEText(html_body, "html"))

    for attachment in mail_config.get("attachments", []):
        if "path" in attachment:
            file_path = attachment["path"]
            with open(file_path, "rb") as f:
                file_data = f.read()
            filename = os.path.basename(file_path)
            mime_type, _ = mimetypes.guess_type(file_path)
            if not mime_type:
                mime_type = "application/octet-stream"
            subtype = mime_type.split("/")[-1]
        elif "data" in attachment and "filename" in attachment:
            file_data = attachment["data"]
            filename = attachment["filename"]
            mime_type = attachment.get("mime_type", "application/octet-stream")
            subtype = mime_type.split("/")[-1]
        else:
            raise ValueError("Invalid attachment format. Must provide 'path' or ('data', 'filename').")

        part = MIMEApplication(file_data, _subtype=subtype)
        part.add_header("Content-Disposition", "attachment", filename=filename)
        message.attach(part)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as smtp:
            smtp.starttls()
            smtp.login(sender, password)
            smtp.sendmail(sender, receivers + cc, message.as_string())
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")



def main():
    failed_invoice_rec_list = [
        {
            "claimNumber": "19227616",
            "invoiceNumber": "INV1059283",
            "status": "VALIDATION_FAILED",
            "statusDescription": "Required field is missing , Data type not matching for column date",
        },
        {
            "claimNumber": "19227634",
            "invoiceNumber": "INV1059267",
            "status": "PROCESSING_FAILED",
            "statusDescription": "No Data found for the claims",
        }
    ]

    file_related_detail = {
        "status": "PROCESSING_COMPLETED",
        "fileName": "SG Weekly Invoices 6112025 (AllowWheel).xlsx",
        "uploadedBy": "testuser",
        "uploadedDate": "2025-06-11T00:00:00",
        "errorRecCount": 2,
        "totalRecCount": 10,
        "pendingRecCount": 0,
        "successRecCount": 0,
        "statusDescription": "FILE PROCESSED SUCCESSFULLY"
    }

    # Generate Excel file binary
    xlsx_data = generate_error_file(failed_invoice_rec_list)

    # Save to disk (optional)
    output_file_path = "error_report.xlsx"
    save_xlsx_file(output_file_path, xlsx_data)

    # Generate email HTML content
    html_body = generate_invoice_email_body(file_related_detail, "error_report.xlsx")

    # Email configuration dictionary
    mail_config = {
        "sender_email": "ikandasa@sgintl.com",
        "sender_password": "IN@Office390",
        "receiver_emails": ["hitesh.paliwal@thinknsolutions.com", "hiteshpaliwal1703@gmail.com"],
        "cc_emails": ["hiteshpaliwal.j@gmail.com"],
        "subject": f'Invoice Process detail for "{file_related_detail["fileName"]}"',
        "html_body": html_body,
        "attachments": [
            {
                "data": xlsx_data,
                "filename": "error_report.xlsx",
                "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
            # Add more attachments here if needed
        ]
    }

    # Send email
    send_mail_notification(mail_config)


if __name__ == "__main__":
    main()
