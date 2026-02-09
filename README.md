```
import os
import shutil
import time
import re
import xlwings as xw


MASTER_PATH = r"C:\FULL\PATH\Master.xlsx"

PREP_FOLDER = r"C:\FULL\PATH\01_prepared_files"
CALC_FOLDER = r"C:\FULL\PATH\02_calculated_files"
VALUE_FOLDER = r"C:\FULL\PATH\03_value_files"

ENTITIES = [
    "APAC",
    "EMEA",
    "AMERICAS",
    "INDIA",
    "UK"
]

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL = "F1"

SHEETS_TO_KEEP = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]


def safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())
