import io
import time
import concurrent.futures
from pathlib import Path
from typing import Optional, List, Dict, Any, Callable, Union

from PIL import Image
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

# HEIC/HEIF support (iPhone photos)
try:
    import pillow_heif  # type: ignore
    pillow_heif.register_heif_opener()
except ImportError:
    pass

MODEL_NAME = "gemini-2.5-flash"

EXTRACTION_PROMPT = """You are an expert OCR assistant. Analyze this handwritten form image carefully.

Extract the following fields from the form:
1. "Name (Chinese)" - The person's name in Chinese characters. Be extremely careful to read the exact characters.
2. "Name (English)" - The person's name in English/romanized form. MUST BE ALL UPPERCASE.
3. "Class" - The class/grade information
4. "Contact Number" - The phone/contact number
5. "Food Type" - The food type/preference

Rules:
1. If a field is empty or you cannot read it, set the value to "" (empty string)
2. Be as accurate as possible with the handwritten text
3. For "Name (English)", convert to ALL UPPERCASE.
4. For "Food Type", if NORMAL return "N", if VEGETARIAN return "VEGE", if HALAL return "HALAL".
5. For "Contact Number", remove all dashes and spaces. If the number starts with "0", replace the leading "0" with "60" (e.g., "012-345 6789" becomes "60123456789"). Ensure you double-check digits carefully.
"""

class ExtractedForm(BaseModel):
    name_chinese: str = Field(alias="Name (Chinese)", default="")
    name_english: str = Field(alias="Name (English)", default="")
    student_class: str = Field(alias="Class", default="")
    contact_number: str = Field(alias="Contact Number", default="")
    food_type: str = Field(alias="Food Type", default="")
    
    model_config = {"extra": "ignore"}

def prepare_image(image_path: Union[str, Path]) -> Image.Image:
    """
    Opens an image and prepares it for Gemini API.
    Specifically handles HEIC conversion by saving to JPEG in memory.
    """
    img = Image.open(image_path)
    
    # Convert HEIC (or images with transparency) to RGB
    if img.mode != "RGB":
        img = img.convert("RGB")
    
    # Save to a BytesIO object as JPEG to ensure compatibility with Gemini SDK
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG')
    img_byte_arr.seek(0)
    
    # Return a new PIL image from the JPEG bytes
    return Image.open(img_byte_arr)

def extract_single_form(
    client: genai.Client, 
    filepath: Union[str, Path], 
    model_name: str = MODEL_NAME,
    cancel_event=None
) -> Dict[str, Any]:
    """
    Extracts data from a single form image.
    Returns a dictionary matching the expected structure.
    """
    filepath_str = Path(filepath).name
    
    # Check cancellation before starting
    if cancel_event and cancel_event.is_set():
        return {
            "_source_file": filepath_str,
            "_status": "error",
            "_error": "Cancelled"
        }

    max_retries = 3
    for attempt in range(max_retries):
        try:
            # Check cancellation
            if cancel_event and cancel_event.is_set():
                return {
                    "_source_file": filepath_str,
                    "_status": "error",
                    "_error": "Cancelled"
                }

            img = prepare_image(filepath)
            
            response = client.models.generate_content(
                model=model_name,
                contents=[EXTRACTION_PROMPT, img],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ExtractedForm,
                ),
            )
            
            # response.parsed will contain the ExtractedForm pydantic model instance
            parsed_data = response.parsed
            
            if parsed_data:
                # Use model_dump(by_alias=True) to get the expected keys like "Name (Chinese)"
                data = parsed_data.model_dump(by_alias=True)
            else:
                raise ValueError("API returned empty parsed data")
                
            data["_source_file"] = filepath_str
            data["_status"] = "success"
            return data

        except Exception as e:
            error_msg = str(e)
            
            if cancel_event and cancel_event.is_set():
                return {
                    "_source_file": filepath_str,
                    "_status": "error",
                    "_error": "Cancelled"
                }

            # If it's a 503 Service Unavailable or 429 Too Many Requests, wait and retry
            if ("503" in error_msg or "429" in error_msg) and attempt < max_retries - 1:
                time.sleep(2 * (attempt + 1))
                continue
                
            return {
                "_source_file": filepath_str,
                "_status": "error",
                "_error": error_msg
            }
    
    return {
        "_source_file": filepath_str,
        "_status": "error",
        "_error": "Max retries exceeded"
    }

def process_images_parallel(
    client: genai.Client,
    filepaths: List[Union[str, Path]],
    model_name: str = MODEL_NAME,
    progress_callback: Optional[Callable[[int, int, str, Dict[str, Any]], None]] = None,
    cancel_event=None,
    max_workers: int = 3
) -> List[Dict[str, Any]]:
    """
    Process multiple images in parallel using ThreadPoolExecutor.
    progress_callback(current_completed, total_files, filename, result_data)
    """
    total = len(filepaths)
    all_results = [None] * total
    completed = 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all tasks and track their original index
        future_to_index = {
            executor.submit(extract_single_form, client, filepath, model_name, cancel_event): i
            for i, filepath in enumerate(filepaths)
        }
        
        for future in concurrent.futures.as_completed(future_to_index):
            idx = future_to_index[future]
            filepath = filepaths[idx]
            filename = Path(filepath).name
            
            try:
                result = future.result()
            except Exception as exc:
                result = {
                    "_source_file": filename,
                    "_status": "error",
                    "_error": str(exc)
                }
            
            all_results[idx] = result
            completed += 1
            
            if progress_callback:
                progress_callback(completed, total, filename, result)
                
    return all_results
