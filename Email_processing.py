import io
import smtplib
from datetime import datetime
from typing import List, Any
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


def send_mail_notification(
    sender_mail_id: str,
    receiver_mail_id: List[str],
    cc_mail_id: List[str],
    file_related_detail: dict,
    attachment_data: Any,
    attachment_filename: str,
    sender_password : str,
    smtp_server: str = "smtp.office365.com",
    smtp_port: int = 587


):
    """
    Sends email with table in body and XLSX attachment via Outlook SMTP.
    """
    subject = f'Invoice Process detail for "{file_related_detail["fileName"]}"'

    uploaded_date = file_related_detail.get("uploadedDate", "").split("T")[0]

    html_table = f"""
    <p>Please find the invoice file processing result as below:</p>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse;">
        <tr><td><strong>File Name</strong></td><td>{file_related_detail["fileName"]}</td></tr>
        <tr><td><strong>Uploaded By</strong></td><td>{file_related_detail["uploadedBy"]}</td></tr>
        <tr><td><strong>Uploaded Date</strong></td><td>{uploaded_date}</td></tr>
        <tr><td><strong>Total Records</strong></td><td>{file_related_detail["totalRecCount"]}</td></tr>
        <tr><td><strong>Error Records</strong></td><td>{file_related_detail["errorRecCount"]}</td></tr>
        <tr><td><strong>Success Records</strong></td><td>{file_related_detail["successRecCount"]}</td></tr>
        <tr><td><strong>Pending Records</strong></td><td>{file_related_detail["pendingRecCount"]}</td></tr>
        <tr><td><strong>Status</strong></td><td>{file_related_detail["status"]}</td></tr>
        <tr><td><strong>Status Description</strong></td><td>{file_related_detail["statusDescription"]}</td></tr>
    </table>
    <p>Please find the error invoice record details in attached spreadsheet: <strong>{attachment_filename}</strong></p>
    <br>
    <p>Thanks,<br>VCIM SUPPORT</p>
    """

    message = MIMEMultipart()
    message['From'] = sender_mail_id
    message['To'] = ", ".join(receiver_mail_id)
    message['Cc'] = ", ".join(cc_mail_id)
    message['Subject'] = subject
    message.attach(MIMEText(html_table, 'html'))

    attachment = MIMEApplication(attachment_data, _subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    attachment.add_header('Content-Disposition', 'attachment', filename=attachment_filename)
    message.attach(attachment)

    all_recipients = receiver_mail_id + cc_mail_id

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as smtp:
            smtp.starttls()
            smtp.login(sender_mail_id, sender_password)
            smtp.sendmail(sender_mail_id, all_recipients, message.as_string())
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

    # Send Email
    send_mail_notification(
        sender_mail_id="ikandasa@sgintl.com",
        receiver_mail_id=["hitesh.paliwal@thinknsolutions.com", "hiteshpaliwal1703@gmail.com"],
        cc_mail_id=["surya.prakash@thinknsolutions.com","shafeeq.kadhar@thinknsolutions.com"],
        file_related_detail=file_related_detail,
        attachment_data=xlsx_data,
        attachment_filename="error_report.xlsx",
        sender_password="IN@Office389"
    )


if __name__ == "__main__":
    main()
