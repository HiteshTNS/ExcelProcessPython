import pandas as pd
import logging
import sys

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
    try:
        # Read Excel
        logger.info(f"Reading Excel file: {file_path}")
        df = pd.read_excel(file_path, engine='openpyxl')

        required_cols = ['VIN', 'Term', 'Start Date', 'Price']

        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            logger.error(f"Missing required columns in file: {missing_cols}")
            return

        # Sort by VIN ascending, then Price descending
        logger.info("Sorting by VIN (A-Z) and Price (Largest to Smallest)")
        df = df.sort_values(by=['VIN', 'Price'], ascending=[True, False]).reset_index(drop=True)

        # Add Count column: 1 where current row matches previous row on VIN, Price, Prev Program, and Submitted By == 1.6
        logger.info("Calculating 'Count' column")
        df['Count'] = (
            (df['VIN'] == df['VIN'].shift(1)) &
            (df['Term'] == df['Term'].shift(1)) &
            (df['Start Date'] == df['Start Date'].shift(1)) &
            (df['Price'] == 1.6)
        ).astype(int)

        # Filter where Count == 1
        filtered_df = df[df['Count'] == 1]
        logger.info(f"Filtered rows where Count == 1: {len(filtered_df)}")

        # Extract unique VINs from filtered rows
        vins = filtered_df['VIN']
        # logger.info(f"Unique VINs extracted: {len(vins)}")

        # Write VINs to new Excel file
        logger.info(f"Writing VINs to: {output_vin_path}")
        vins.to_frame().to_excel(output_vin_path, index=False, engine='openpyxl')

        logger.info("Processing completed successfully.")

    except Exception as e:
        logger.error(f"Error processing file: {e}", exc_info=True)

if __name__ == "__main__":
    # Change these paths as needed
    input_file = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\Invoice_Detail_SGD202502.xlsx"
    output_file = r"C:\Users\hitesh.paliwal\Desktop\ExcelProject\filtered_vins.xlsx"
    process_excel(input_file, output_file)
