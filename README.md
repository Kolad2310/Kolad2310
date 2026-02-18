```

# ============================= trend_engine.py =============================

from logger import logger

def compute_trends(df):
    """
    Compute rolling metrics with higher sensitivity
    suitable for trending financial data (millions).
    """

    df = df.sort_values("Date")

    # Use 15-day rolling window for better sensitivity
    df["MA"] = df["Exposure"].rolling(15, min_periods=1).mean()

    df["Volatility"] = df["Exposure"].rolling(15, min_periods=1).std()

    logger.info("Trend metrics computed (15-day window).")

    return df

# ============================= anomaly.py =============================

from logger import logger

def detect_anomaly(df):
    """
    Hybrid anomaly detection:
    1. Z-score threshold (statistical)
    2. Percentage drop threshold (business shock)
    3. Absolute drop threshold (large value drop)
    """

    # Safe Z-score calculation
    df["Z_score"] = (
        (df["Exposure"] - df["MA"]) /
        df["Volatility"].replace(0, 1)
    )

    # Percentage change
    df["Pct_Change"] = df["Exposure"].pct_change()

    # Absolute drop from moving average
    df["Abs_Drop"] = df["MA"] - df["Exposure"]

    df["Alert"] = df.apply(
        lambda row: "YES" if (
            abs(row["Z_score"]) > 1.5          # statistical
            or row["Pct_Change"] < -0.10       # >10% drop
            or row["Abs_Drop"] > 400000        # >4 lakh drop
        ) else "NO",
        axis=1
    )

    logger.info("Hybrid anomaly detection completed.")

    return df

df = pd.read_excel(file_path)

validate_schema(df)

# Clean date
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df = df.dropna(subset=["Date"])

# Clean exposure
df["Exposure"] = pd.to_numeric(df["Exposure"], errors="coerce")
df = df.dropna(subset=["Exposure"])



# ============================= chart_generator.py =============================

import os
import matplotlib.pyplot as plt
from config import PROCESSED_FOLDER
from logger import logger


def generate_trend_chart(df):
    """
    Generates trend chart for latest processed data.
    Keeps only one latest PNG file.
    """

    try:
        # Ensure output folder exists
        os.makedirs(PROCESSED_FOLDER, exist_ok=True)

        chart_path = os.path.join(PROCESSED_FOLDER, "trend_latest.png")

        # Clear old chart if exists
        if os.path.exists(chart_path):
            os.remove(chart_path)

        plt.figure(figsize=(12, 6))

        # Plot Exposure
        plt.plot(df["Date"], df["Exposure"], label="Exposure")

        # Plot Moving Average
        plt.plot(df["Date"], df["MA"], label="Moving Average (15D)")

        # Highlight Alert Points
        alerts = df[df["Alert"] == "YES"]
        if not alerts.empty:
            plt.scatter(
                alerts["Date"],
                alerts["Exposure"],
                marker="o",
                s=100,
                label="Alert"
            )

        plt.title("Risk Exposure Trend")
        plt.xlabel("Date")
        plt.ylabel("Exposure")
        plt.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()

        plt.savefig(chart_path)
        plt.close()

        logger.info("Trend chart generated successfully.")

    except Exception as e:
        logger.error(f"Chart generation failed: {str(e)}")




from chart_generator import generate_trend_chart


df = detect_anomaly(df)

# Generate chart
generate_trend_chart(df)

