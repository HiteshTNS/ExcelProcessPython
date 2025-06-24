import json
import pandas as pd
import logging
import sys
import time
import os

# Setup logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler("excel_processing.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger()

# File processing configurations
PROCESSING_CONFIGS = {
    "default": { #For First Filtration
        "required_columns": ['VIN', 'TERM', 'START DATE', 'PRICE'],
        "sort_by": [('VIN', True), ('PRICE', False)],
        "count_formula": "filtration_for_file_1",
        "export_columns": 'ALL'
    },
    "custom_file_2": { #For second Filtration
        "required_columns": ['FORM', 'VIN', 'PURE RISK TYPE'],
        "sort_by": [('FORM', True), ('VIN', True), ('PURE RISK TYPE', True)],
        "count_formula": "filtration_for_file_2",
        "export_columns": 'ALL'
    }
}

def read_input_file(file_path: str) -> pd.DataFrame:
    try:
        if not os.path.exists(file_path):
            logger.error(f"Input file not found: {file_path}")
            raise FileNotFoundError(f"Input file does not exist at the specified path:\n  → {file_path}")

        file_extension = os.path.splitext(file_path)[1].lower()
        if file_extension == ".xlsx":
            logger.info("Detected Excel (.xlsx) file format.")
            data_frame = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        elif file_extension == ".csv":
            logger.info("Detected CSV (.csv) file format.")
            data_frame = pd.read_csv(file_path, dtype=str)
        else:
            raise ValueError("Unsupported file format. Only .xlsx and .csv are supported.")

        original_columns = {col.strip().upper(): col.strip() for col in data_frame.columns}
        data_frame.rename(columns={value: key for key, value in original_columns.items()}, inplace=True)
        data_frame.attrs['original_columns'] = original_columns

        for col in data_frame.select_dtypes(include='object').columns:
            data_frame[col] = data_frame[col].str.strip()

        return data_frame

    except FileNotFoundError as fnf:
        raise fnf
    except Exception as e:
        logger.error(f"Error reading input file: {e}", exc_info=True)
        raise

def validate_columns(data_frame: pd.DataFrame, required_columns: list):
    missing_cols = [col for col in required_columns if col not in data_frame.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")

def preprocess_data(data_frame: pd.DataFrame, sort_by: list[tuple[str, bool]]) -> pd.DataFrame:
    try:
        if 'PRICE' in data_frame.columns:
            data_frame['PRICE'] = pd.to_numeric(data_frame['PRICE'], errors='coerce')
        sort_columns = [col for col, _ in sort_by]
        sort_orders = [asc for _, asc in sort_by]
        data_frame.sort_values(by=sort_columns, ascending=sort_orders, kind='mergesort', inplace=True)
        data_frame.reset_index(drop=True, inplace=True)
        return data_frame
    except Exception as e:
        logger.error("Error during preprocessing: %s", e, exc_info=True)
        raise

def calculate_count_column(data_frame: pd.DataFrame) -> pd.DataFrame:
    try:
        data_frame['COUNT'] = (
            (data_frame['VIN'] == data_frame['VIN'].shift(1)) &
            (data_frame['TERM'] == data_frame['TERM'].shift(1)) &
            (data_frame['START DATE'] == data_frame['START DATE'].shift(1)) &
            (data_frame['PRICE'] == 1.6)
        ).astype(int)
        return data_frame
    except Exception as e:
        logger.error(f"Error calculating 'COUNT' column: {e}", exc_info=True)
        raise

def calculate_count_custom_e2(data_frame: pd.DataFrame) -> pd.DataFrame:
    try:
        if not all(col in data_frame.columns for col in ['FORM', 'VIN', 'PURE RISK TYPE']):
            raise ValueError("One or more required columns are missing for custom E2 formula.")
        pure_risk_type_cleaned = data_frame['PURE RISK TYPE'].str.upper()
        data_frame['COUNT'] = (
            (data_frame['FORM'] == data_frame['FORM'].shift(1)) &
            (data_frame['VIN'] == data_frame['VIN'].shift(1)) &
            (pure_risk_type_cleaned.shift(1) == "KEY") &
            (pure_risk_type_cleaned == "ROADSIDE")
        ).astype(int)
        return data_frame
    except Exception as e:
        logger.error("Error in custom E2 formula: %s", e, exc_info=True)
        raise

def apply_count_formula(data_frame: pd.DataFrame, formula_key: str) -> pd.DataFrame:
    if formula_key == "filtration_for_file_1":
        return calculate_count_column(data_frame)
    elif formula_key == "filtration_for_file_2":
        return calculate_count_custom_e2(data_frame)
    else:
        raise ValueError(f"Unknown formula: {formula_key}")

def filter_vins_with_count_one(data_frame: pd.DataFrame, export_columns):
    count = 0 # variable to store the number of filtered vin's
    try:
        if 'COUNT' not in data_frame.columns:
            raise ValueError("COUNT column not found.")

        filtered_df = data_frame[data_frame['COUNT'] == 1]

        if export_columns == 'ALL':
            final_df = filtered_df.copy()
        else:
            final_df = filtered_df[export_columns].copy()

        # Restore original column casing before output
        original_columns = data_frame.attrs.get('original_columns', {})
        final_df.rename(columns=original_columns, inplace=True)

        record_list = final_df.to_dict(orient='records')
        count = len(record_list)
        logger.info(f"Filtered rows where Count == 1: {count}")
        return final_df, record_list

    except Exception as e:
        logger.error("Error filtering rows with COUNT == 1: %s", e, exc_info=True)
        raise

def export_dataframe(data_frame: pd.DataFrame, output_path: str):
    try:
        if output_path.lower().endswith('.xlsx'):
            data_frame.to_excel(output_path, index=False, engine='openpyxl')
        elif output_path.lower().endswith('.csv'):
            data_frame.to_csv(output_path, index=False)
        else:
            raise ValueError("Unsupported file format. Use '.xlsx' or '.csv'.")
        logger.info(f"Data exported successfully to: {output_path}")
    except PermissionError:
        logger.error(f"Permission denied while writing to '{output_path}'.")
        logger.error("Please make sure the file is **not open in Excel** or any other application.")
        raise
    except Exception as e:
        logger.error(f"Error exporting DataFrame: {e}", exc_info=True)
        raise

def process_excel(file_path: str, output_file_path: str, config_key: str = "default"):
    global filtered_records_file1, filtered_records_file2
    start_time = time.time()
    try:
        config = PROCESSING_CONFIGS.get(config_key)
        if not config:
            logger.error(f"Unknown processing config: {config_key}")
            return

        data_frame = read_input_file(file_path)
        validate_columns(data_frame, config['required_columns'])
        data_frame = preprocess_data(data_frame, config['sort_by'])
        data_frame = apply_count_formula(data_frame, config['count_formula'])

        filtered_df, filtered_list = filter_vins_with_count_one(data_frame, config.get('export_columns', 'ALL'))
        export_dataframe(filtered_df, output_file_path)

        if config_key == 'default':
            filtered_records_file1 = filtered_list
        elif config_key == 'custom_file_2':
            filtered_records_file2 = filtered_list

        logger.info(f"Processing completed in {round(time.time() - start_time, 2)} seconds.")

    except PermissionError:
        logger.error("Processing aborted due to file access error.")
    except ValueError as ve:
        logger.error(f"Processing aborted: {ve}")
    except Exception as e:
        logger.error("Unexpected error during processing: %s", e, exc_info=True)

if __name__ == "__main__":
    # process_excel(
    #     file_path=r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\Invoice_Detail_SGD202502.xlsx",
    #     output_file_path=r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\filtered_vins1.xlsx",
    #     config_key="default"
    # )

    # process_excel(
    #     file_path=r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\VAS Roadside dupes 2025-05-06T1409.csv",
    #     output_file_path=r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\filtered_vins2.csv",
    #     config_key="custom_file_2"
    # )

    # print("\nFiltered Records for File 1 (Wrapped JSON):")
    # print(json.dumps({"filtered_data_file1": filtered_records_file1}, indent=2))

    print("\nFiltered Records for File 2 (Wrapped JSON):")
    print(json.dumps({"filtered_data_file_2": filtered_records_file2}, indent=2))
