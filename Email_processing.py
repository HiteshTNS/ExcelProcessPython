import io
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

def main():
    # Sample failed invoice records
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

    # Generate XLSX Binary Data
    xlsx_data = generate_error_file(failed_invoice_rec_list)

    # Save to file
    output_file_path = "error_report.xlsx"
    save_xlsx_file(output_file_path, xlsx_data)

if __name__ == "__main__":
    main()
