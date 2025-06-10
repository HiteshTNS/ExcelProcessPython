import pandas as pd
import json
import os
import chardet
import logging

# Setup logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler("CPR_Invoice_Logger.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

def detect_encoding(file_path, num_bytes=10000):
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read(num_bytes)
        result = chardet.detect(raw_data)
        logger.info(f"Detected file encoding: {result['encoding']}")
        return result['encoding']
    except Exception as e:
        logger.error(f"Failed to detect file encoding: {e}")
        raise

def read_file(file_path, encoding):
    try:
        _, ext = os.path.splitext(file_path)
        logger.info(f"Reading file: {file_path}")

        if ext.lower() == '.csv':
            df = pd.read_csv(file_path, encoding=encoding, delimiter=',', engine='python')
        elif ext.lower() in ['.xls', '.xlsx']:
            if ext == '.xlsx':
                df = pd.read_excel(file_path, engine='openpyxl')
            elif ext == '.xls':
                df = pd.read_excel(file_path, engine='xlrd')
            else:
                raise ValueError("Unsupported Excel file format.")
        else:
            raise ValueError("Unsupported file format. Use .csv or .xlsx")

        df.columns = df.columns.str.strip()
        logger.info("File read successfully with normalized headers.")
        return df
    except Exception as e:
        logger.error(f"Error reading file: {e}")
        raise

def removespecialcharacters(val):
        return int(str(val).replace('$', '').replace(',', '').strip())

def clean_value(val):
    if pd.isna(val) or val is None:
        return "Null"
    try:
        num = float(str(val).replace('$', '').replace(',', '').strip())
        return str(int(num)) if num.is_integer() else str(num)
    except:
        return str(val).strip()

def build_record(row):
    try:
        return {
            "claimNumber": clean_value(row.get("Claim #")),
            "amount": clean_value(row.get("Amount")),
            "recJson": {
                "fileName": clean_value(row.get("File #")),
                "claimNumber": clean_value(row.get("Claim #")),
                "contractNumber": clean_value(row.get("Contact")),
                "insured": clean_value(row.get("Insured")),
                "stateOfLoss": clean_value(row.get("State of Loss")),
                "businessType": clean_value(row.get("Business Type")),
                "txn": clean_value(row.get("Txn #")),
                "type": clean_value(row.get("Type")),
                "date": clean_value(row.get("Date")),
                "amount": clean_value(row.get("Amount")),
                "0to30Days": clean_value(row.get("0-30 days")),
                "31to60Days": clean_value(row.get("31-60 days")),
                "61to90Days": clean_value(row.get("61-90 days")),
                "91+Days": clean_value(row.get("91+ days")),
                "netBalance": clean_value(row.get("Net Balance")),
                "mileage": clean_value(row.get("Mileage"))
            }
        }
    except Exception as e:
        logger.error(f"Error building record for row: {e}")
        raise

def process_file(file_path, encoding):
    try:
        df = read_file(file_path, encoding)
        df = df.fillna("Null")

        standard_fee = []
        non_standard_fee = []

        logger.info("Processing rows...")

        for index, row in df.iterrows():
            try:
                claim_val = row.get("Claim #", "Null")
                contact_val = row.get("Contact", "Null")
                insured_val = row.get("Insured", "Null")

                # Abort if critical fields are missing
                if (pd.isna(claim_val) or str(claim_val).strip() == "Null") and \
                   (pd.isna(contact_val) or str(contact_val).strip() == "Null") and \
                   (pd.isna(insured_val) or str(insured_val).strip() == "Null"):
                    logger.error(f"Row {index + 2}: Claim, Contact, and Insured all missing. Aborting.")
                    raise ValueError("Required fields missing")

                amount_val = removespecialcharacters(row.get("Amount", "Null"))
                record = build_record(row)

                if amount_val in [85, 140]:
                    standard_fee.append(record)
                else:
                    non_standard_fee.append(record)
            except Exception as row_err:
                logger.warning(f"Skipping row {index + 2} due to error: {row_err}")
                continue

        logger.info("Finished processing file.")
        return {
            "standardFee": standard_fee,
            "nonStandardFee": non_standard_fee
        }

    except Exception as e:
        logger.error(f"Error in processing file: {e}")
        raise

if __name__ == "__main__":
    try:
        file_path = r'C:\Users\hitesh.paliwal\Desktop\ExcelProject\CPR Invoice\Input\CPR Insurance template - April(Sheet1).csv'
        output_path = r'C:\Users\hitesh.paliwal\Desktop\ExcelProject\CPR Invoice\Output\output_fee_data.json'

        encoding = detect_encoding(file_path)
        result = process_file(file_path, encoding)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        with open(output_path, "w") as file:
            json.dump(result, file, indent=4, default=str)

        logger.info(f"Output written to: {output_path}")
        logger.info(f"Standard Fee Records: {len(result['standardFee'])}")
        logger.info(f"Non-Standard Fee Records: {len(result['nonStandardFee'])}")
    except Exception as final_error:
        logger.critical(f"Script failed: {final_error}")
