import os
import logging
from fastapi import FastAPI, UploadFile, File
from typing import List, Dict, Any
from collections import defaultdict
from pydantic import BaseModel
from unstructured.partition.pptx import partition_pptx
from unstructured.cleaners.core import clean, clean_extra_whitespace
from unstructured.cleaners.extract import  (
    extract_email_address,
    extract_ip_address,
    extract_us_phone_number,
    extract_datetimetz,
    extract_ordered_bullets
)

from rapidfuzz import fuzz

# Initialize FastAPI app and logging
app = FastAPI()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("pptx_structure_microservice")


def looks_like_table(text_block: str) -> bool:
    """
    Detect if a block of text resembles a table structure.

    Args:
        text_block (str): A string block to be analyzed.

    Returns:
        bool: True if text resembles a table, False otherwise.
    """
    logger.debug("Checking if text block looks like a table")
    lines = text_block.strip().splitlines()
    count_pipes = sum(1 for line in lines if '|' in line)
    count_tabs = sum(1 for line in lines if '\t' in line)
    return len(lines) >= 2 and (count_pipes > 0 or count_tabs > 0)


def label_text_block(text_block: str, raw_category: str) -> str:
    """
    Classify a block of text based on its content or metadata.

    Args:
        text_block (str): The text to classify.
        raw_category (str): Initial category from metadata.

    Returns:
        str: A refined content category.
    """
    logger.debug("Labeling text block with category logic")
    if looks_like_table(text_block):
        return "LikelyTable"
    if extract_email_address(text_block):
        return "Email"
    if extract_ip_address(text_block):
        return "IPAddress"
    if extract_us_phone_number(text_block):
        return "PhoneNumber"
    if extract_datetimetz(text_block):
        return "DateTime"
    if any(extract_ordered_bullets(text_block)):
        return "OrderedListItem"
    return raw_category if raw_category != "UncategorizedText" else "Text"


@app.post("/extract_structure")
async def extract_pptx_structure(file: UploadFile = File(...)) -> Dict[str, Any]:
    """
    Extract and categorize structured content from a PowerPoint (PPTX) file.

    Args:
        file (UploadFile): The uploaded PPTX file.

    Returns:
        dict: A structured representation of the slides and categorized elements.
    """
    logger.info(f"Received file: {file.filename}")

    # Save the uploaded file to a temporary location
    temp_path = f"/tmp/{file.filename}"
    try:
        with open(temp_path, "wb") as temp_file:
            temp_file.write(await file.read())
        logger.debug(f"Saved file to {temp_path}")
    except Exception as file_error:
        logger.exception("Failed to save uploaded PPTX file")
        return {"success": False, "error": str(file_error)}

    try:
        # Extract structured elements using Unstructured.io
        pptx_elements = partition_pptx(filename=temp_path, infer_table_structure=True)
        logger.info(f"Parsed {len(pptx_elements)} elements from the PPTX file")

        slides_by_number: Dict[int, List[Dict[str, Any]]] = defaultdict(list)

        for element in pptx_elements:
            text_content = (element.text or "").strip()
            if not text_content:
                continue

            metadata = element.metadata.to_dict()
            slide_number = int(metadata.get("page_number", 0) or 0)
            base_category = element.category or "Unknown"
            labeled_type = label_text_block(text_content, base_category)

            slides_by_number[slide_number].append({
                "type": labeled_type,
                "text": text_content,
                "metadata": {
                    "slide_number": slide_number,
                    "filename": metadata.get("filename"),
                    "coordinates": metadata.get("coordinates"),
                    "element_id": metadata.get("element_id"),
                    "raw_category": base_category,
                }
            })

        # Convert to sorted list format
        structured_slides: List[Dict[str, Any]] = [
            {"slide": slide_num, "elements": slides_by_number[slide_num]}
            for slide_num in sorted(slides_by_number)
        ]

        logger.info("Extraction and labeling completed successfully")
        return {"success": True, "slides": structured_slides}

    except Exception as processing_error:
        logger.exception("Failed to extract structure from PPTX")
        return {"success": False, "error": str(processing_error)}
    finally:
        # Clean up the temporary file
        try:
            os.remove(temp_path)
            logger.debug(f"Temporary file {temp_path} deleted")
        except Exception as cleanup_error:
            logger.warning(f"Could not delete temp file: {cleanup_error}")
