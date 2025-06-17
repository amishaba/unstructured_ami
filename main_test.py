# from unstructured.partition.auto import partition
# from unstructured.staging.base import convert_to_dataframe

# # 1. Partition your PDF file to get a list of elements
# elements = partition("nvidia.pdf", include_page_breaks=True)

# # 2. Convert the list of elements into a pandas DataFrame
# df = convert_to_dataframe(elements)

# # 3. Select only the desired columns for your spreadsheet
# # This step is optional, but helps create a cleaner table.
# df = df[["text", "type", "page_number", "filename"]]

# # 4. Save the DataFrame to a CSV file
# df.to_csv("Q4_Report_elements.csv", index=False)

# print("Successfully created Q4_Report_elements.csv")




# from unstructured.partition.pdf import partition_pdf
# from unstructured.cleaners.core import clean_extra_whitespace
# from unstructured.chunking.title import chunk_by_title
# from unstructured.staging.base import elements_to_json
# import json

# # Load the PDF and partition it
# elements = partition_pdf(
#     filename="nvidia.pdf",
#     include_metadata=True,
#     # infer_table_structure=True,  # Triggers OCR or PDF layout processing if needed
# )

# print(f"Extracted {len(elements)} elements")
# print(f"First 3 element types: {[type(el).__name__ for el in elements[:3]]}")

# # Optional: Clean each element's text
# for el in elements:
#     if hasattr(el, 'text') and el.text:
#         el.text = clean_extra_whitespace(el.text)

# # Chunk by structure (headings, paragraphs)
# chunks = chunk_by_title(elements, max_characters=500)

# import pandas as pd

# data = []
# for el in chunks:
#     data.append({
#         "type": type(el).__name__,
#         "text": el.text,
#         "metadata": el.metadata.to_dict() if el.metadata else {},
#     })

# # Create DataFrame
# df = pd.DataFrame(data)

# # Export to Excel
# df.to_excel("nvidia_output.xlsx", index=False)

# print("✅ Excel output saved to nvidia_output.xlsx")

# print(f"\nChunked into {len(chunks)} sections")
# print(f"First chunk preview:\n{chunks[0].text[:300]}...")

# # Serialize to JSON
# json_output = elements_to_json(chunks)
# with open("nvidia_output.json", "w", encoding="utf-8") as f:
#     json.dump(json_output, f, indent=2, ensure_ascii=False)

# print("\n✅ Processing complete. Output saved to nvidia_output.json")










# import os
# import camelot
# import pandas as pd
# from unstructured.partition.pdf import partition_pdf

# # Optional: prepend Poppler path if needed
# poppler_path = r"C:\Users\nitin\OneDrive\Desktop\intern project\unstructredio\poppler-24.08.0\Library\bin"
# os.environ["PATH"] = poppler_path + ";" + os.environ["PATH"]
# os.environ["UNSTRUCTURED_DISABLE_OCR"] = "true"

# # --- UNSTRUCTURED: Parse structure ---
# elements = partition_pdf(
#     filename="nvidia.pdf",
#     include_metadata=True,
#     infer_table_structure=True,  # this uses layout model but not OCR
#     extract_images_in_pdf=False,
#     ocr_languages="",            # disables pytesseract fallback
#     hi_res_model_name="yolox",   # or None if you don’t need layout model
#     strategy="fast"              # speeds up + disables OCR fallback
# )

# # Print a few native Unstructured Elements
# print("\n🔍 Unstructured Elements Preview:")
# for el in elements[:10]:
#     print(f"- {type(el).__name__}: {el.text[:100]}")

# # --- CAMELOT: Extract tables ---
# import camelot
# import pandas as pd

# # --- Camelot Table Extraction ---
# print("\n📊 Extracting tables using Camelot...")
# tables = camelot.read_pdf("nvidia.pdf", pages="all", flavor="stream")  # try "stream" for text-aligned PDFs

# if tables and tables.n > 0:
#     with pd.ExcelWriter("nvidia_tables.xlsx", engine="openpyxl") as writer:
#         for i, table in enumerate(tables):
#             df = table.df
#             df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)
#     print(f"✅ Saved {tables.n} tables to nvidia_tables.xlsx")
# else:
#     print("⚠️ No tables found with Camelot. Excel file not created.")



import os
import camelot
import pandas as pd
from unstructured.partition.pdf import partition_pdf

# --- Setup Paths ---
pdf_path = "nvidia.pdf"
excel_path = "nvidia_report.xlsx"
poppler_path = r"C:\Users\nitin\OneDrive\Desktop\intern project\unstructredio\poppler-24.08.0\Library\bin"

os.environ["PATH"] = poppler_path + ";" + os.environ["PATH"]
os.environ["UNSTRUCTURED_DISABLE_OCR"] = "true"

# --- UNSTRUCTURED: Parse structure ---
elements = partition_pdf(
    filename=pdf_path,
    include_metadata=True,
    infer_table_structure=True,
    extract_images_in_pdf=False,
    ocr_languages="",
    hi_res_model_name="yolox",
    strategy="fast"
)

# --- Step 1: Extract Unstructured Elements into DataFrame ---
structured_data = []
for el in elements:
    structured_data.append({
        "Type": type(el).__name__,
        "Category": el.category,
        "Text": el.text.strip()[:500],  # Preview text
        "Page": getattr(el.metadata, "page_number", None)
    })
df_structure = pd.DataFrame(structured_data)

# --- Step 2: Extract Tables using Camelot ---
print("\n📊 Extracting tables using Camelot...")
tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")

# --- Step 3: Write Everything to One Beautiful Excel File ---
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    # Cover/overview sheet
    pd.DataFrame({
        "Report Title": ["NVIDIA Earnings Report Summary"],
        "Total Elements Extracted": [len(elements)],
        "Total Tables Found": [tables.n]
    }).to_excel(writer, sheet_name="Summary", index=False)

    # Document structure
    df_structure.to_excel(writer, sheet_name="Document Structure", index=False)

    # Each table to its own tab
    if tables.n > 0:
        for i, table in enumerate(tables):
            table.df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)
    else:
        pd.DataFrame({"Message": ["No tables detected by Camelot."]}).to_excel(writer, sheet_name="Tables", index=False)

print(f"✅ Final Excel report saved as: {excel_path}")
