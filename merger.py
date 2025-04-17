# merger.py
import os
import zipfile
import pandas as pd
from pathlib import Path
from datetime import datetime

EXTRACT_DIR = "extracted"
OUTPUT_FILE = f"merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

def extract_zip(zip_path):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(EXTRACT_DIR)

def merge_excel_files():
    merged_data = []
    header_set = None

    for file in Path(EXTRACT_DIR).glob("*.xlsx"):
        try:
            df = pd.read_excel(file)
            if header_set is None:
                header_set = df.columns.tolist()
            elif df.columns.tolist() != header_set:
                continue  # Skip file with different headers
            merged_data.append(df)
        except Exception as e:
            print(f"Error reading {file.name}: {e}")

    if merged_data:
        result = pd.concat(merged_data, ignore_index=True)
        result.to_excel(OUTPUT_FILE, index=False)
        return OUTPUT_FILE
    else:
        return None
