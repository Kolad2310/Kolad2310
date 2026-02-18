```
# ============================= logger.py =============================

"""
Central Logging Configuration

Features:
- Auto-creates logs directory
- Writes logs to file
- Prints logs to console
- Prevents duplicate handlers
- Production-ready format
"""

import os
import logging
from config import LOG_FILE


# --------------------------------------------------
# Ensure logs directory exists
# --------------------------------------------------
log_directory = os.path.dirname(LOG_FILE)
os.makedirs(log_directory, exist_ok=True)


# --------------------------------------------------
# Create Logger
# --------------------------------------------------
logger = logging.getLogger("RiskTrendEngine")
logger.setLevel(logging.INFO)


# Prevent duplicate handlers if re-imported
if not logger.handlers:

    # ---------------- File Handler ----------------
    file_handler = logging.FileHandler(LOG_FILE)
    file_handler.setLevel(logging.INFO)

    # ---------------- Console Handler ----------------
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # ---------------- Log Format ----------------
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    )

    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # ---------------- Add Handlers ----------------
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)


# Optional: Reduce noise from matplotlib
logging.getLogger("matplotlib").setLevel(logging.WARNING)


# --------------------------------------------------
# Example Usage (for testing)
# --------------------------------------------------
if __name__ == "__main__":
    logger.info("Logger initialized successfully.")
    logger.warning("This is a warning example.")
    logger.error("This is an error example.")
