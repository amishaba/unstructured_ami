import os
import logging
from fastapi import FastAPI, UploadFile, File
from typing import List, Dict, Any
from collections import defaultdict
from unstructured.partition.pptx import partition_pptx
from unstructured.cleaners.core import clean, clean_extra_whitespace
from unstructured.cleaners.extract import  (
    extract_email_address,
    extract_ip_address,
    extract_us_phone_number,
    extract_datetimetz,
    extract_ordered_bullets
)


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




async def save_uploaded_file(file: UploadFile, destination_path: str) -> None:
    """
    Save an uploaded file to disk.

    Args:
        file (UploadFile): Uploaded file from FastAPI.
        destination_path (str): Full path to write file to.
    """
    try:
        with open(destination_path, "wb") as temp_file:
            temp_file.write(await file.read())
        logger.debug(f"File saved to {destination_path}")
    except Exception as e:
        logger.exception("Error saving file to disk")
        raise e

def process_elements_by_slide(pptx_elements) -> List[Dict[str, Any]]:
    """
    Organize and label extracted elements by slide number.

    Args:
        pptx_elements: List of elements extracted from the PPTX.

    Returns:
        List[Dict]: Structured list of slides and elements.
    """
    slides_by_number: Dict[int, List[Dict[str, Any]]] = defaultdict(list)

    for element in pptx_elements:
        text_content = (element.text or "").strip()
        if not text_content:
            continue

        metadata = element.metadata.to_dict()
        slide_number = int(metadata.get("page_number", 0) or 0)
        base_category = element.category or "Unknown"
        labeled_type = label_text_block(text_content, base_category)

        # Add the element to its corresponding slide in the dictionary
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

    # Convert dictionary to a sorted list format for consistent API output
    structured_slides = [
        {"slide": slide_num, "elements": slides_by_number[slide_num]}
        for slide_num in sorted(slides_by_number)
    ]

    return structured_slides


@app.post("/extract_structure")
async def extract_pptx_structure(file: UploadFile = File(...)) -> Dict[str, Any]:
    """
    API endpoint to extract and categorize structured content from a PowerPoint (PPTX) file.

    Args:
        file (UploadFile): The uploaded PPTX file.

    Returns:
        dict: A structured representation of slides and categorized elements.
    """
    logger.info(f"Received file: {file.filename}")
    temp_path = f"/tmp/{file.filename}"

    try:
        # Step 1: Save uploaded file to disk
        await save_uploaded_file(file, temp_path)

        # Step 2: Parse the PPTX into structured elements
        pptx_elements = partition_pptx(filename=temp_path, infer_table_structure=True)
        logger.info(f"Parsed {len(pptx_elements)} elements from the PPTX file")

        # Step 3: Organize and label elements by slide
        structured_slides = process_elements_by_slide(pptx_elements)

        logger.info("Extraction and labeling completed successfully")
        return {"success": True, "slides": structured_slides}

    except Exception as e:
        logger.exception("Failed to extract structure from PPTX")
        return {"success": False, "error": str(e)}

    finally:
        # Step 4: Clean up temporary file
        try:
            os.remove(temp_path)
            logger.debug(f"Temporary file {temp_path} deleted")
        except Exception as cleanup_error:
            logger.warning(f"Could not delete temp file: {cleanup_error}")
