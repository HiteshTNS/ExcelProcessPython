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

REQUIRED_COLUMNS = ['VIN', 'Term', 'Start Date', 'Price']


def read_input_file(file_path: str) -> pd.DataFrame:
    try:
        file_extension = os.path.splitext(file_path)[1].lower()
        if file_extension == ".xlsx":
            logger.info("Detected Excel (.xlsx) file format.")
            data_frame = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        elif file_extension == ".csv":
            logger.info("Detected CSV (.csv) file format.")
            data_frame = pd.read_csv(file_path, dtype=str)
        else:
            raise ValueError("Unsupported file format. Only .xlsx and .csv are supported.")

        data_frame.columns = [col.strip() for col in data_frame.columns]  # Clean column names
        return data_frame
    except Exception as e:
        logger.error(f"Error reading input file: {e}", exc_info=True)
        raise


def validate_columns(data_frame: pd.DataFrame):
    try:
        missing_cols = [col for col in REQUIRED_COLUMNS if col not in data_frame.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns: {missing_cols}")
    except Exception as e:
        logger.error(f"Column validation failed: {e}", exc_info=True)
        raise


def preprocess_data(data_frame: pd.DataFrame) -> pd.DataFrame:
    try:
        data_frame['Price'] = pd.to_numeric(data_frame['Price'], errors='coerce')
        data_frame.sort_values(by=['VIN', 'Price'], ascending=[True, False], kind='mergesort', inplace=True)
        data_frame.reset_index(drop=True, inplace=True)
        return data_frame
    except Exception as e:
        logger.error(f"Error during preprocessing: {e}", exc_info=True)
        raise


def calculate_count_column(data_frame: pd.DataFrame) -> pd.DataFrame:
    try:
        data_frame['Count'] = (
            (data_frame['VIN'] == data_frame['VIN'].shift(1)) &
            (data_frame['Term'] == data_frame['Term'].shift(1)) &
            (data_frame['Start Date'] == data_frame['Start Date'].shift(1)) &
            (data_frame['Price'] == 1.6)
        ).astype(int)
        return data_frame
    except Exception as e:
        logger.error(f"Error calculating 'Count' column: {e}", exc_info=True)
        raise


def filter_vins_with_count_one(data_frame: pd.DataFrame) -> pd.DataFrame:
    try:
        filtered_df = data_frame[data_frame['Count'] == 1][['VIN']]
        logger.info(f"Filtered rows where Count == 1: {len(filtered_df)}")
        return filtered_df
    except Exception as e:
        logger.error(f"Error filtering VINs: {e}", exc_info=True)
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
    except Exception as e:
        logger.error(f"Error exporting DataFrame: {e}", exc_info=True)
        raise


def process_excel(file_path: str, output_file_path: str):
    start_time = time.time()
    try:
        data_frame = read_input_file(file_path)
        validate_columns(data_frame)
        data_frame = preprocess_data(data_frame)
        data_frame = calculate_count_column(data_frame)
        filtered_vins = filter_vins_with_count_one(data_frame)
        export_dataframe(filtered_vins, output_file_path)
        logger.info(f"Processing completed in {round(time.time() - start_time, 2)} seconds.")
    except Exception as e:
        logger.error(f"Processing failed: {e}", exc_info=True)


if __name__ == "__main__":
    input_file = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\Invoice_Detail_SGD202502.xlsx"
    output_file = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\filtered_vins.csv"  # Can also use .csv
    process_excel(input_file, output_file)
