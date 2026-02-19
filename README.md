```
# ============================= ingestion.py =============================

import os
import pandas as pd

from trend_engine import compute_trends
from anomaly import detect_anomaly
from chart_generator import generate_trend_chart
from config import INPUT_FOLDER, PROCESSED_FOLDER
from logger import logger


def process_risk_file(file_path):
    """
    Consolidates ALL Excel files inside input folder
    and generates single consolidated output + multi-line chart.
    """

    try:
        logger.info("Starting consolidated processing...")

        # --------------------------------------------------
        # 1️⃣ Delete old processed files
        # --------------------------------------------------
        if os.path.exists(PROCESSED_FOLDER):
            for file in os.listdir(PROCESSED_FOLDER):
                os.remove(os.path.join(PROCESSED_FOLDER, file))

        os.makedirs(PROCESSED_FOLDER, exist_ok=True)

        # --------------------------------------------------
        # 2️⃣ Read all Excel files from input folder
        # --------------------------------------------------
        all_data = []

        for file in os.listdir(INPUT_FOLDER):
            if file.endswith(".xlsx"):
                full_path = os.path.join(INPUT_FOLDER, file)

                df = pd.read_excel(full_path)

                # Clean
                df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
                df["Exposure"] = pd.to_numeric(df["Exposure"], errors="coerce")
                df = df.dropna(subset=["Date", "Exposure"])

                df = df.sort_values("Date")

                # Add source column
                df["Source_File"] = file

                all_data.append(df)

        if not all_data:
            logger.warning("No Excel files found in input folder.")
            return

        # --------------------------------------------------
        # 3️⃣ Consolidate
        # --------------------------------------------------
        consolidated_df = pd.concat(all_data, ignore_index=True)

        # --------------------------------------------------
        # 4️⃣ Compute trends per file separately
        # --------------------------------------------------
        final_df_list = []

        for file_name, group in consolidated_df.groupby("Source_File"):
            group = compute_trends(group)
            group = detect_anomaly(group)
            final_df_list.append(group)

        final_df = pd.concat(final_df_list)

        # --------------------------------------------------
        # 5️⃣ Save consolidated Excel
        # --------------------------------------------------
        output_path = os.path.join(
            PROCESSED_FOLDER,
            "consolidated_output.xlsx"
        )

        final_df.to_excel(output_path, index=False)

        logger.info("Consolidated Excel saved.")

        # --------------------------------------------------
        # 6️⃣ Generate Multi-Line Chart
        # --------------------------------------------------
        generate_trend_chart(final_df)

        logger.info("Processing completed successfully.\n")

    except Exception as e:
        logger.error(f"Error during consolidated processing: {str(e)}")
        print("❌ Error occurred. Check logs.")










# ============================= chart_generator.py =============================

import os
import matplotlib.pyplot as plt
from config import PROCESSED_FOLDER
from logger import logger


def generate_trend_chart(df):
    """
    Generates consolidated multi-line trend chart.
    One line per source file.
    Keeps only latest PNG.
    """

    try:
        os.makedirs(PROCESSED_FOLDER, exist_ok=True)

        chart_path = os.path.join(PROCESSED_FOLDER, "trend_latest.png")

        # Delete old PNG if exists
        if os.path.exists(chart_path):
            os.remove(chart_path)

        plt.figure(figsize=(12, 6))

        # Plot one line per file
        for file_name, group in df.groupby("Source_File"):
            plt.plot(
                group["Date"],
                group["Exposure"],
                label=file_name
            )

        plt.title("Consolidated Risk Exposure Trend")
        plt.xlabel("Date")
        plt.ylabel("Exposure")
        plt.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()

        plt.savefig(chart_path)
        plt.close()

        logger.info("Multi-line trend chart generated successfully.")

    except Exception as e:
        logger.error(f"Chart generation failed: {str(e)}")
