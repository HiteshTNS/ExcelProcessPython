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
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".xlsx":
            logger.info("Detected Excel (.xlsx) file format.")
            df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        elif ext == ".csv":
            logger.info("Detected CSV (.csv) file format.")
            df = pd.read_csv(file_path, dtype=str)
        else:
            raise ValueError("Unsupported file format. Only .xlsx and .csv are supported.")

        df.columns = [col.strip() for col in df.columns]  # Clean column names
        return df
    except Exception as e:
        logger.error(f"Error reading input file: {e}", exc_info=True)
        raise


def validate_columns(df: pd.DataFrame):
    try:
        missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns: {missing_cols}")
    except Exception as e:
        logger.error(f"Column validation failed: {e}", exc_info=True)
        raise


def preprocess_data(df: pd.DataFrame) -> pd.DataFrame:
    try:
        df['Price'] = pd.to_numeric(df['Price'], errors='coerce')
        # df['Term'] = df['Term'].astype(str)
        df.sort_values(by=['VIN', 'Price'], ascending=[True, False], kind='mergesort', inplace=True)
        df.reset_index(drop=True, inplace=True)
        return df
    except Exception as e:
        logger.error(f"Error during preprocessing: {e}", exc_info=True)
        raise


def calculate_count_column(df: pd.DataFrame) -> pd.DataFrame:
    try:
        df['Count'] = (
            (df['VIN'] == df['VIN'].shift(1)) &
            (df['Term'] == df['Term'].shift(1)) &
            (df['Start Date'] == df['Start Date'].shift(1)) &
            (df['Price'] == 1.6)
        ).astype(int)
        return df
    except Exception as e:
        logger.error(f"Error calculating 'Count' column: {e}", exc_info=True)
        raise


def filter_and_export_vins(df: pd.DataFrame, output_path: str):
    try:
        filtered_df = df[df['Count'] == 1]
        logger.info(f"Filtered rows where Count == 1: {len(filtered_df)}")
        vins = filtered_df[['VIN']]
        vins.to_excel(output_path, index=False, engine='openpyxl')
        logger.info(f"Filtered VINs written to: {output_path}")
    except Exception as e:
        logger.error(f"Error writing VINs to file: {e}", exc_info=True)
        raise


def process_excel(file_path: str, output_vin_path: str):
    start_time = time.time()
    try:
        df = read_input_file(file_path)
        validate_columns(df)
        df = preprocess_data(df)
        df = calculate_count_column(df)
        filter_and_export_vins(df, output_vin_path)
        logger.info(f"Processing completed in {round(time.time() - start_time, 2)} seconds.")
    except Exception as e:
        logger.error(f"Processing failed: {e}", exc_info=True)


if __name__ == "__main__":
    input_file = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\Invoice_Detail_SGD202502.csv"
    output_file = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\filtered_vins.xlsx"
    process_excel(input_file, output_file)
