import pandas as pd
import logging
import sys
import time

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

def process_excel(file_path: str, output_vin_path: str):
    start_time = time.time()

    try:
        logger.info(f"Reading Excel file: {file_path}")
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)

        required_cols = ['VIN', 'Term', 'Start Date', 'Price']
        df.columns = [col.strip() for col in df.columns]  # Strip whitespace from column names
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            logger.error(f"Missing required columns in file: {missing_cols}")
            return

        # Convert numeric column
        df['Price'] = pd.to_numeric(df['Price'], errors='coerce')
        df['Term'] = df['Term'].astype(str)

        logger.info("Sorting by VIN (A-Z) and Price (Largest to Smallest)")
        df.sort_values(by=['VIN', 'Price'], ascending=[True, False], kind='mergesort', inplace=True)
        df.reset_index(drop=True, inplace=True)

        logger.info("Calculating 'Count' column (countifs-like)")
        df['Count'] = (
            (df['VIN'] == df['VIN'].shift(1)) &
            (df['Term'] == df['Term'].shift(1)) &
            (df['Start Date'] == df['Start Date'].shift(1)) &
            (df['Price'] == 1.6)
        ).astype(int)

        filtered_df = df[df['Count'] == 1]
        logger.info(f"Filtered rows where Count == 1: {len(filtered_df)}")

        vins = filtered_df[['VIN']]
        logger.info(f"Writing VINs to: {output_vin_path}")
        vins.to_excel(output_vin_path, index=False, engine='openpyxl')

        logger.info("Processing completed successfully.")
        logger.info(f"Total time taken: {round(time.time() - start_time, 2)} seconds")

    except Exception as e:
        logger.error(f"Error processing file: {e}", exc_info=True)

if __name__ == "__main__":
    input_file = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\Invoice_Detail_SGD202502.xlsx"
    output_file = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\filtered_vins.xlsx"
    process_excel(input_file, output_file)
