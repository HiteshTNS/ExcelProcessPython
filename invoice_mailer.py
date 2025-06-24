import os
import json
from excel_processing import  process_excel
from Email_processing import  send_mail_notification

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
        <tr><td><strong>Success Records</strong></td><td>{file_detail["successRecCount"]}</td></tr>
        <tr><td><strong>Pending Records</strong></td><td>{file_detail["pendingRecCount"]}</td></tr>
        <tr><td><strong>Status</strong></td><td>{file_detail["status"]}</td></tr>
        <tr><td><strong>Status Description</strong></td><td>{file_detail["statusDescription"]}</td></tr>
    </table>
    <p>Please find the error invoice record details in attached spreadsheet: <strong>{attachment_name}</strong></p>
    <br>
    <p>Thanks,<br>VCIM SUPPORT</p>
    """



def main():
    # File paths
    input_file1 = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\Invoice_Detail_SGD202502.xlsx"
    output_file1 = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\filtered_vins1.xlsx"

    input_file2 = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\VAS Roadside dupes 2025-05-06T1409.csv"
    output_file2 = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\filtered_vins2.csv"

    # Step 1: Run Excel processing
    process_excel(
        file_path=input_file1,
        output_file_path=output_file1,
        # config_key="custom_file_2"
    )

    process_excel(
        file_path=input_file2,
        output_file_path=output_file2,
        config_key="custom_file_2"
    )

    # Step 2: Prepare email details
    file_detail = {
        "status": "PROCESSING_COMPLETED",
        "fileName": os.path.basename(input_file1),
        "uploadedBy": "automated-script",
        "uploadedDate": "2025-06-20T00:00:00",
        "totalRecCount": "N/A",  # If you want to count original rows, read it in
        "pendingRecCount": 0,
        "successRecCount": 0,
        "statusDescription": "Processed with Python automation"
    }

    html_body = generate_invoice_email_body(file_detail, os.path.basename(output_file1))

    mail_config = {
        "sender_email": "ikandasa@sgintl.com",
        "sender_password": "IN@Office390",
        "receiver_emails": ["hitesh.paliwal@thinknsolutions.com"],
        "cc_emails": ["hiteshpaliwal.j@gmail.com"],
        "subject": f"Processed VAS File - {file_detail['fileName']}",
        "html_body": html_body,
        "attachments": [
            {"path":output_file1},
            {"path":output_file2}
        ]
    }

    # Step 3: Send email
    send_mail_notification(mail_config)

if __name__ == "__main__":
    main()
