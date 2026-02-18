```
# ============================= ingestion.py =============================

"""
Central Processing Pipeline

This file:
1. Reads Excel input file
2. Validates schema
3. Cleans Date and Exposure columns
4. Computes trend metrics
5. Detects anomalies (hybrid logic)
6. Saves processed Excel file
7. Saves data to SQLite DB
8. Generates latest trend chart PNG
"""

import os
import pandas as pd

from validator import validate_schema
from trend_engine import compute_trends
from anomaly import detect_anomaly
from database import save_to_database
from chart_generator import generate_trend_chart
from config import PROCESSED_FOLDER
from logger import logger


def process_risk_file(file_path):
    """
    Main ingestion pipeline triggered by Watchdog
    """

    try:
        logger.info(f"Processing file started: {file_path}")

        # --------------------------------------------------
        # 1️⃣ Read Excel file
        # --------------------------------------------------
        df = pd.read_excel(file_path)

        # --------------------------------------------------
        # 2️⃣ Validate schema (must contain Date & Exposure)
        # --------------------------------------------------
        validate_schema(df)

        # --------------------------------------------------
        # 3️⃣ Clean & Standardize Data
        # --------------------------------------------------

        # Convert Date column safely
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

        # Convert Exposure to numeric
        df["Exposure"] = pd.to_numeric(df["Exposure"], errors="coerce")

        # Remove invalid rows
        df = df.dropna(subset=["Date", "Exposure"])

        # Sort properly
        df = df.sort_values("Date")

        logger.info("Data cleaning completed.")

        # --------------------------------------------------
        # 4️⃣ Compute Trend Metrics
        # --------------------------------------------------
        df = compute_trends(df)

        # --------------------------------------------------
        # 5️⃣ Detect Anomalies (Hybrid Logic)
        # --------------------------------------------------
        df = detect_anomaly(df)

        # --------------------------------------------------
        # 6️⃣ Save Processed Excel File
        # --------------------------------------------------
        os.makedirs(PROCESSED_FOLDER, exist_ok=True)

        output_path = os.path.join(
            PROCESSED_FOLDER,
            "processed_" + os.path.basename(file_path)
        )

        df.to_excel(output_path, index=False)

        logger.info(f"Processed file saved at: {output_path}")

        # --------------------------------------------------
        # 7️⃣ Save to SQLite Database
        # (Auto-creates DB if not present)
        # --------------------------------------------------
        try:
            save_to_database(df)
        except Exception as db_error:
            logger.warning(f"Database save skipped: {str(db_error)}")

        # --------------------------------------------------
        # 8️⃣ Generate Trend Chart (Keeps only latest PNG)
        # --------------------------------------------------
        try:
            generate_trend_chart(df)
        except Exception as chart_error:
            logger.warning(f"Chart generation skipped: {str(chart_error)}")

        logger.info("Processing completed successfully.\n")

    except Exception as e:
        logger.error(f"Error processing file: {str(e)}")
        print("❌ Error occurred. Check logs.")
