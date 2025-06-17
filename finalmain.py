import os
import camelot
import pandas as pd
from unstructured.partition.pdf import partition_pdf
from unstructured.cleaners.core import clean_extra_whitespace
from unstructured.chunking.title import chunk_by_title
from unstructured.staging.base import elements_to_json

# --- ENVIRONMENT SETUP ---
pdf_path = "nvidia.pdf"
excel_path = "nvidia_full_report.xlsx"
poppler_path = r"C:\Users\nitin\OneDrive\Desktop\intern project\unstructredio\poppler-24.08.0\Library\bin"

os.environ["PATH"] = poppler_path + ";" + os.environ["PATH"]
os.environ["UNSTRUCTURED_DISABLE_OCR"] = "true"

# --- STEP 1: PARTITION PDF USING UNSTRUCTURED ---
elements = partition_pdf(
    filename=pdf_path,
    include_metadata=True,
    infer_table_structure=True,
    extract_images_in_pdf=False,
    ocr_languages="",
    hi_res_model_name="yolox",
    strategy="fast"
)

# --- STEP 2: CLEAN TEXT ---
for el in elements:
    if hasattr(el, "text") and el.text:
        el.text = clean_extra_whitespace(el.text)

# --- STEP 3: CHUNKING ---
chunks = chunk_by_title(elements, max_characters=600)

# --- STEP 4: STAGING TO JSON ---
staged_json = elements_to_json(chunks)

# --- STEP 5: CAMELOT TABLE EXTRACTION ---
print("\n📊 Extracting tables with Camelot...")
tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")

# --- STEP 6: EXPORT TO EXCEL ---
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    # Summary tab
    pd.DataFrame({
        "Report": ["NVIDIA Earnings Summary"],
        "Total Raw Elements": [len(elements)],
        "Chunked Sections": [len(chunks)],
        "Tables Found": [tables.n]
    }).to_excel(writer, sheet_name="Summary", index=False)

    # Unstructured Elements
    pd.DataFrame([
        {
            "Type": type(el).__name__,
            "Category": el.category,
            "Text": el.text[:500],
            "Page": getattr(el.metadata, "page_number", None)
        }
        for el in elements
    ]).to_excel(writer, sheet_name="Raw Elements", index=False)

    # Chunked output
    # --- Chunked output ---
    pd.DataFrame([
        {
            "Chunk Type": type(c).__name__,
            "Text": c.text[:1000],
            "Section Title": c.metadata.to_dict().get("section_title", None) if c.metadata else None
        }
        for c in chunks
    ]).to_excel(writer, sheet_name="Chunks", index=False)



    # Camelot tables
    if tables.n > 0:
        for i, table in enumerate(tables):
            table.df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)
    else:
        pd.DataFrame({"Message": ["No tables detected."]}).to_excel(writer, sheet_name="Tables", index=False)

print(f"✅ Report saved: {excel_path}")
