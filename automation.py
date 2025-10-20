import os
import json
import glob
import smtplib
import logging
from datetime import datetime
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
import pandas as pd

# ------------------------
# Config & Paths
# ------------------------
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
LOG_DIR = ROOT / "logs"
CONFIG_PATH = ROOT / "config" / "email_settings.json"

# ------------------------
# Logging Setup
# ------------------------
LOG_DIR.mkdir(parents=True, exist_ok=True)
log_file = LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ------------------------
# Helpers
# ------------------------

def load_config(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Missing config file: {path}")
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    return cfg

def ensure_folders():
    for p in (DATA_DIR, OUTPUT_DIR, LOG_DIR):
        p.mkdir(parents=True, exist_ok=True)
    logger.info("Folders checked/created: data, output, logs")

def list_excel_files() -> list[Path]:
    files = [Path(p) for p in glob.glob(str(DATA_DIR / "*.xlsx"))]
    logger.info("Found %s Excel files in /data", len(files))
    return files

def merge_excels(files: list[Path]) -> pd.DataFrame:
    if not files:
        logger.warning("No Excel files found to merge.")
        return pd.DataFrame()

    frames = []
    for fp in files:
        try:
            df = pd.read_excel(fp)
            df["__source_file"] = fp.name
            frames.append(df)
            logger.info("Loaded %s -> %s", fp.name, df.shape)
        except Exception as e:
            logger.exception("Failed to read %s: %s", fp, e)

    if not frames:
        return pd.DataFrame()

    merged = pd.concat(frames, ignore_index=True)
    logger.info("Merged shape: %s", merged.shape)
    return merged

# ------------------------
# Main Execution
# ------------------------

if __name__ == "__main__":
    ensure_folders()
    files = list_excel_files()
    merged_df = merge_excels(files)

    if not merged_df.empty:
        output_path = OUTPUT_DIR / f"merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        merged_df.to_excel(output_path, index=False)
        logger.info("Saved merged Excel file to: %s", output_path)
    else:
        logger.info("No data merged. Please add .xlsx files to the /data folder.")
