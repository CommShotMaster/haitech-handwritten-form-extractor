# 🖊️ Handwritten Form Extractor

Extract handwritten text from form images using **Gemini 2.5 Flash** (free API) and save the results to Excel.

## Quick Start

### 1. Get a Free Gemini API Key

1. Go to [Google AI Studio](https://aistudio.google.com/apikey)
2. Click **"Create API Key"**
3. Copy the key

### 2. Configure API Key

Create a `.env` file in this folder (copy from `.env.example`):

```
GEMINI_API_KEY=your_actual_api_key_here
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Add Your Form Images

Place your handwritten form images in the **`uploads/`** folder.

Supported formats: `.jpg`, `.jpeg`, `.png`, `.bmp`, `.webp`, `.heic`, `.heif`

### 5. Run the Extractor

```bash
python extract_form.py
```

### 6. Check the Output

The extracted data will be saved as an Excel file in the **`output/`** folder.

## How It Works

1. The script scans the `uploads/` folder for form images
2. Each image is sent to **Gemini 2.5 Flash** with a prompt to extract all handwritten fields
3. Gemini returns structured JSON with field names and values
4. All results are combined into a single Excel file (one row per form)

## Folder Structure

```
uploads/    → Place your handwritten form images here
output/     → Extracted Excel files appear here
```
