"""
Handwritten Form Extractor using Gemini 2.5 Flash
Extracts handwritten text from form images and saves to Excel.
Uses the new google.genai SDK.
"""

import os
import sys
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from PIL import Image
from dotenv import load_dotenv
from google import genai
from google.genai import types

# HEIC/HEIF support (iPhone photos)
try:
    import pillow_heif  # type: ignore
    pillow_heif.register_heif_opener()
except ImportError:
    pass


# ─── Configuration ───────────────────────────────────────────────────────────

UPLOADS_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR = Path(__file__).parent / "output"
SUPPORTED_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".webp", ".heic", ".heif")
MODEL_NAME = "gemini-2.5-flash"

# Use the shared extractor core
from extractor_core import process_images_parallel


def setup_api():
    """Configure the Gemini API with the user's key."""
    load_dotenv()
    api_key = os.getenv("GEMINI_API_KEY")

    if not api_key or api_key == "your_api_key_here":
        print("=" * 60)
        print("ERROR: Gemini API key not found!")
        print()
        print("Steps to fix:")
        print("1. Go to https://aistudio.google.com/apikey")
        print("2. Create a free API key")
        print("3. Create a .env file in this folder with:")
        print("   GEMINI_API_KEY=your_actual_key_here")
        print()
        print("(You can copy .env.example as a template)")
        print("=" * 60)
        sys.exit(1)

    client = genai.Client(api_key=api_key)
    return client


def get_image_files():
    """Get all supported image files from the uploads directory."""
    if not UPLOADS_DIR.exists():
        UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

    image_files = []
    for ext in SUPPORTED_EXTENSIONS:
        image_files.extend(UPLOADS_DIR.glob(f"*{ext}"))
        image_files.extend(UPLOADS_DIR.glob(f"*{ext.upper()}"))

    # Remove duplicates (in case of case-insensitive filesystem)
    seen = set()
    unique_files = []
    for f in image_files:
        if f.resolve() not in seen:
            seen.add(f.resolve())
            unique_files.append(f)

    return sorted(unique_files)


def cli_progress_callback(completed: int, total: int, filename: str, data: dict):
    if data.get("_status") == "success":
        print(f"  ✅ [{completed}/{total}] Extracted {len(data) - 2} fields from {filename}")
    else:
        err = data.get("_error", "Unknown error")
        print(f"  ❌ [{completed}/{total}] Error processing {filename}: {err}")


def save_to_excel(all_data, output_path):
    """Save extracted data to an Excel file."""
    df = pd.DataFrame(all_data)

    # Move _source_file to the first column
    cols = df.columns.tolist()
    if "_source_file" in cols:
        cols.remove("_source_file")
        cols = ["_source_file"] + cols
    df = df[cols]

    # Rename for readability
    df = df.rename(columns={"_source_file": "Source File"})

    # Remove internal columns from display if no errors
    if "_error" in df.columns and df["_error"].isna().all():
        df = df.drop(columns=["_error"])
    if "_raw_text" in df.columns and df["_raw_text"].isna().all():
        df = df.drop(columns=["_raw_text"])

    df.to_excel(output_path, index=False, engine="openpyxl")
    return df


def main():
    print("=" * 60)
    print("🖊️  Handwritten Form Extractor (Gemini 2.5 Flash)")
    print("=" * 60)
    print()

    # Setup
    client = setup_api()

    # Find images
    image_files = get_image_files()

    if not image_files:
        print(f"📁 No images found in: {UPLOADS_DIR.resolve()}")
        print()
        print("Place your handwritten form images in the 'uploads' folder.")
        print(f"Supported formats: {', '.join(SUPPORTED_EXTENSIONS)}")
        sys.exit(0)

    print(f"📁 Found {len(image_files)} image(s) in uploads/")
    print()

    print("🚀 Starting parallel extraction...")
    all_data = process_images_parallel(
        client=client, 
        filepaths=image_files, 
        progress_callback=cli_progress_callback,
        max_workers=3
    )

    # Save to Excel
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = OUTPUT_DIR / f"extracted_forms_{timestamp}.xlsx"

    df = save_to_excel(all_data, output_path)

    print()
    print("=" * 60)
    print(f"✅ Done! Processed {len(image_files)} form(s)")
    print(f"📊 Excel saved to: {output_path.resolve()}")
    print()
    print("Preview:")
    print(df.to_string(index=False))
    print("=" * 60)


if __name__ == "__main__":
    main()
