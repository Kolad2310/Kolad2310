```
✅ Watchdog listener
✅ Validation
✅ Ingestion
✅ Trend analysis (Moving Average + Volatility)
✅ Anomaly detection (Z-Score)
✅ SQLite storage
✅ Logging
✅ Clear file-style headers (as comments





"""
================================================================================
FILE: risk_trend_engine.py
AUTHOR: Your Name
DESCRIPTION:
Event-driven Risk Trend Monitoring Engine using Python Watchdog.

This script:
1. Monitors a folder for new Excel risk files
2. Validates schema
3. Computes trend metrics (30-day moving average, volatility)
4. Performs anomaly detection (Z-score)
5. Stores results in SQLite database
6. Logs all activity
================================================================================
"""

# ============================== IMPORT SECTION ==============================

import os
import time
import logging
import sqlite3
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ============================== CONFIG SECTION ==============================

"""
Central configuration block.
Modify paths here only.
"""

INPUT_FOLDER = r"data/input"
PROCESSED_FOLDER = r"data/processed"
DATABASE_NAME = "risk_trend.db"
LOG_FILE = "risk_engine.log"

# Create folders if not present
os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# ============================== LOGGER SETUP ==============================

"""
Production-grade logging.
All events/errors get logged into a file.
"""

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

# ============================== VALIDATION MODULE ==============================

def validate_schema(df):
    """
    Ensures required columns exist.
    Prevents downstream errors due to schema mismatch.
    """
    required_cols = ["Date", "Exposure"]

    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")

    logging.info("Schema validation successful.")
    return True

# ============================== TREND ENGINE ==============================

def compute_trends(df):
    """
    Computes rolling financial metrics:
    - 30-day Moving Average
    - Rolling Volatility (Std Dev)
    """

    df = df.sort_values("Date")

    # Rolling 30-day average exposure
    df["30D_MA"] = df["Exposure"].rolling(30).mean()

    # Rolling volatility (risk measure)
    df["Volatility"] = df["Exposure"].rolling(30).std()

    logging.info("Trend metrics computed.")
    return df

# ============================== ANOMALY DETECTION ==============================

def detect_anomaly(df):
    """
    Z-Score anomaly detection:
    Z = (Current Exposure - Moving Average) / Volatility

    If |Z| > 2 → Flag as anomaly.
    """

    df["Z_score"] = (
        (df["Exposure"] - df["30D_MA"]) / df["Volatility"]
    )

    df["Alert"] = df["Z_score"].apply(
        lambda x: "YES" if abs(x) > 2 else "NO"
    )

    logging.info("Anomaly detection completed.")
    return df

# ============================== DATABASE MODULE ==============================

def save_to_database(df):
    """
    Saves processed data into SQLite database.
    Acts as lightweight data warehouse.
    """

    conn = sqlite3.connect(DATABASE_NAME)

    df.to_sql(
        "risk_data",
        conn,
        if_exists="append",
        index=False
    )

    conn.close()

    logging.info("Data saved to SQLite database.")

# ============================== ALERT SYSTEM ==============================

def check_and_alert(df):
    """
    Simple alert mechanism.
    In production, this could send:
    - Email
    - Teams notification
    - Slack message
    """

    alerts = df[df["Alert"] == "YES"]

    if not alerts.empty:
        logging.warning("Anomaly detected in exposure!")
        print("⚠️ Anomaly detected! Check log file.")
    else:
        logging.info("No anomalies detected.")

# ============================== INGESTION PIPELINE ==============================

def process_risk_file(file_path):
    """
    Full processing pipeline triggered by Watchdog event.
    """

    try:
        logging.info(f"Processing file: {file_path}")

        # Step 1: Load Excel file
        df = pd.read_excel(file_path)

        # Step 2: Validate structure
        validate_schema(df)

        # Step 3: Convert Date column properly
        df["Date"] = pd.to_datetime(df["Date"])

        # Step 4: Compute trends
        df = compute_trends(df)

        # Step 5: Detect anomalies
        df = detect_anomaly(df)

        # Step 6: Save processed output
        output_path = os.path.join(
            PROCESSED_FOLDER,
            "processed_" + os.path.basename(file_path)
        )

        df.to_excel(output_path, index=False)

        # Step 7: Store in database
        save_to_database(df)

        # Step 8: Check for alerts
        check_and_alert(df)

        logging.info("File processed successfully.\n")

    except Exception as e:
        logging.error(f"Error processing file: {str(e)}")
        print("❌ Error occurred. Check logs.")

# ============================== WATCHDOG LISTENER ==============================

class RiskFileHandler(FileSystemEventHandler):
    """
    Event handler class.
    Triggers when new file is created in monitored folder.
    """

    def on_created(self, event):
        if not event.is_directory:
            if event.src_path.endswith(".xlsx"):
                print(f"📂 New file detected: {event.src_path}")
                process_risk_file(event.src_path)

# ============================== MAIN EXECUTION ==============================

def start_watching():
    """
    Starts Watchdog observer.
    Runs continuously until manually stopped.
    """

    observer = Observer()
    observer.schedule(
        RiskFileHandler(),
        INPUT_FOLDER,
        recursive=False
    )

    observer.start()

    print("🚀 Risk Trend Engine is running...")
    logging.info("Risk Trend Engine started.")

    try:
        while True:
            time.sleep(5)
    except KeyboardInterrupt:
        observer.stop()
        logging.info("Engine stopped manually.")

    observer.join()


# Entry point
if __name__ == "__main__":
    start_watching()



New Excel File Dropped
        ↓
Watchdog Detects Event
        ↓
Schema Validation
        ↓
Trend Computation (30D MA + Volatility)
        ↓
Z-score Anomaly Detection
        ↓
Save Processed File
        ↓
Store in SQLite
        ↓
Alert if Needed
        ↓
Log Everything


