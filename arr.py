import os
import re
import pandas as pd
from docx import Document
from docx.opc.exceptions import PackageNotFoundError # For handling corrupted docx

# --- Configuration Constants ---
INPUT_FOLDER = "input_docs"
OUTPUT_FILE = "Output_1.xlsx" # All names with categories and total count
OUTPUT_SORTED_FILE = "Output_sorted.xlsx" # All names sorted by category with counts within each category
NEW_CATEGORY_LABEL = "NEW CHECK" # Label for unrecognised categories

# Full category label for long ones
FULL_CUTI_LABEL = "CUTI TAHUN/CPMH/C.SAKIT/C.TAMAT PERKHIDMATAN/C.EHSAN/C.SAMBUNG SAIN/ CUTI PERALIHAN"
PEG_BERTUGAS_LABEL = "PEG BERTUGAS/SJN/KERANI/KOS/KOK/DUTY KERANI"

# Mapping for section headers to the desired CATEGORY output.
# IMPORTANT: Order matters! More specific keys should come BEFORE less specific ones.
SECTION_MAP = {
    # Specific Locations / Penempatan
    "DEPO LOG": "PENEMPATAN LUAR UNIT", # Specific location before general Penempatan
    "ATT PGK EKO COY": "PENEMPATAN DALAM UNIT",
    "E KOMP": "PENEMPATAN DALAM UNIT", # SIG specific
    "PENEMPATAN": "PENEMPATAN DALAM UNIT", # General Penempatan
    "TUGAS TETAP LLP": "PENEMPATAN DALAM UNIT", # Often introduces Penempatan sections
    "PENTADBIRAN": "PENTADBIRAN / REHAT SELEPAS BERTUGAS", # Placed after specific duties/locations

    # Specific Duties / Bertugas
    "JL KERANI": "JURULATIH KERANI",
    "BERTUGAS PEJABAT": "BERTUGAS PEJABAT / TUGAS LUAR BN",
    "BERTUGAS STOR": "BERTUGAS PEJABAT / TUGAS LUAR BN", # CHQ specific
    "BERTUGAS ARMSKOTE": "BERTUGAS PEJABAT / TUGAS LUAR BN", # CHQ specific
    "BERTUGAS": "BERTUGAS PEJABAT / TUGAS LUAR BN", # General Bertugas
    "STORE": "BERTUGAS PEJABAT / TUGAS LUAR BN", # MOR specific (variant spelling of STOR)
    "STOR": "BERTUGAS PEJABAT / TUGAS LUAR BN", # General Stor
    "ARMSKOTE": "BERTUGAS PEJABAT / TUGAS LUAR BN", # General Armskote
    "PEJABAT": "BERTUGAS PEJABAT / TUGAS LUAR BN", # General Pejabat
    "GUARD": "BERTUGAS PEJABAT / TUGAS LUAR BN", # MOR specific
    "KOS": PEG_BERTUGAS_LABEL, # SIG specific
    "KOK": PEG_BERTUGAS_LABEL, # SIG specific
    "DB": PEG_BERTUGAS_LABEL, # PRS specific Duty B
    "DO": PEG_BERTUGAS_LABEL, # PRS specific Duty O
    "DVR/ RO CO": "DRIVER CO & OPSO", # SIG specific

    # Ops / Activities
    "RONDAAN HUTAN": "RONDAAN/RECCE", # PSK specific (more specific)
    "RONDAAN": "RONDAAN/RECCE", # PSK specific (general)
    "PROJEK": "PROJEK", # PRS specific
    "P2B": "P2B", # PSK specific
    "REBRO": "REBRO", # SIG specific
    "OPS ROOM": "OPS ROOMS", # More specific
    "OPS": "OPS ROOMS", # SIG specific (general)

    # Attendance / Presence
    "HADIR BARIS": "HADIR BERBARIS", # More specific
    "HADIR BERBARIS": "HADIR BERBARIS", # Explicit
    "BARIS": "HADIR BERBARIS", # General Baris
    "HADIR": "HADIR BERBARIS", # General Hadir

    # Leave / Absence / Permissions
    "KURSUS PERALIHAN": FULL_CUTI_LABEL, # Specific type of Cuti/Kursus
    "KEBENARAN AKHIR DATANG": "KEBENARAN AKHIR DATANG", # General
    "KEBENARAN AKHER DATANG": "KEBENARAN AKHIR DATANG", # Variant spelling of above
    "KAD": "KEBENARAN AKHIR DATANG", # Abbreviation
    "DATANG": "KEBENARAN AKHIR DATANG", # Keyword
    "KEBENARAN KELUAR": "KEBENARAN KELUAR", # Correct spelling first
    "KEBENARN KELUAR": "KEBENARAN KELUAR", # Variant spelling of above
    "KELUAR": "KEBENARAN KELUAR", # Keyword
    "CPHM": FULL_CUTI_LABEL, # PRS specific Cuti
    "CUTI EHSAN": FULL_CUTI_LABEL, # General
    "CUTI TAHUN": FULL_CUTI_LABEL, 
    "CUTI": FULL_CUTI_LABEL, 

    # Admin / Other Categories
    "REHAT": "PENTADBIRAN / REHAT SELEPAS BERTUGAS", # PSK specific
    "KURSUS": "KURSUS DLM NEGERI", # SIG specific
    "ATT": "ATTCH A/B/C/CHQ/MARKAS BN/ATTCH SBT", # KAP specific Attachment
    "MUSLIM": "NON - MUSLIM", # PGK specific (Verify logic: key MUSLIM -> NON-MUSLIM category?)
    "NON": "NON - MUSLIM",
    "DENTAL": "DENTAL", # PRS specific

    # ADD NEW CATEGORIES HERE
    # "PSU": "PSU CATEGORY",
}

# --- Helps clean the unique name ---
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    cleaned_text = re.sub(r'[\u200c\u200d\ufeff\u2060\u005F]', '', text)     # Remove common invisible characters and underscore first
    cleaned_text = cleaned_text.replace('\u00a0', ' ')     # Replace non-breaking spaces (U+00A0) with regular spaces
    cleaned_text = re.sub(r"[^\w\s]", "", cleaned_text).upper().strip()    # Keep alphanumeric and spaces, remove other punctuation
    return cleaned_text

def parse_personnel_line(text):
    """
    Attempts to parse a line into (number, rank, name).
    Handles CHQ style (e.g., "1.", "01.", "👉 01.") and HSW style (e.g., "-", "*", "👉").
    Rank pattern includes digits (e.g., for PW2).
    Removes common invisible Unicode characters before regex matching. Not sure why some docx have invisible characters.
    """
    text = text.strip()
    if not text:
        return None, None, None

    # Preprocessing: Remove invisible characters and replace non-breaking spaces
    cleaned_text = re.sub(r'[\u200c\u200d\ufeff\u2060]', '', text)
    cleaned_text = cleaned_text.replace('\u00a0', ' ')

    # Rank pattern includes letters, digits, dot, slash
    rank_pattern = r"([a-zA-Z0-9.\/]+)"
    # Personnel Number pattern: 3 to 6 digits
    personnel_num_pattern = r"(\d{3,6})"

    extracted_name = None
    number = None
    rank = None

    # Pattern 1: CHQ style
    match = re.match(r"^\s*(?:[👉]\s*)?(?:\d+\.?\s*)?" + personnel_num_pattern + r"\s+" + rank_pattern + r"\s+(.+)", cleaned_text, re.IGNORECASE)
    if match:
        number, rank, extracted_name = match.group(1), match.group(2), match.group(3).strip()

    # Pattern 2: HSW style (if Pattern 1 failed)
    if not number:
        match = re.match(r"^\s*(?:[-*👉]\s*)?" + personnel_num_pattern + r"\s+" + rank_pattern + r"\s+(.+)", cleaned_text, re.IGNORECASE)
        if match:
            number, rank, extracted_name = match.group(1), match.group(2), match.group(3).strip()

    # Process if either pattern matched
    if number and rank and extracted_name:
        # Remove any trailing content in parentheses
        cleaned_name = re.sub(r"\s*\([^)]*\)\s*$", "", extracted_name).strip()
        return number, rank.upper(), cleaned_name.upper()
    else:
        return None, None, None

# This is where the magic happens which categorises name under the right category
def process_document(filepath, unit_name):
    rows = []
    try:
        doc = Document(filepath)
        current_category = None # Track the category determined by the last valid header

        print(f"Processing: {filepath}")

        for para in doc.paragraphs:
            original_text = para.text
            text_for_headers = original_text.strip()
            if not text_for_headers:
                continue

            norm_text_header = normalize_text(text_for_headers)
            is_header = False # Flag: Did this line match a known header?

            # 1. Check if this line matches a known Section Header
            for header_key, category_value in SECTION_MAP.items():
                normalized_header_key = normalize_text(header_key)
                if not normalized_header_key: continue

                header_matched = False
                # 1st  Check: Does the paragraph START with the header key?
                if norm_text_header.startswith(normalized_header_key):
                    # Relaxed length check for startswith
                    if len(norm_text_header) < len(normalized_header_key) + 35:
                        header_matched = True
                        match_method = "startswith"
                # 2nd Check: If not starts with, check if key is IN text using relaxed heuristics
                elif normalized_header_key in norm_text_header:
                    # Relaxed Heuristic 1: Paragraph length isn't excessively longer than key
                    len_check = len(norm_text_header) < len(normalized_header_key) + 30 
                    # Relaxed Heuristic 2: The non-matching part of the paragraph is relatively short
                    structure_check = len(norm_text_header.replace(normalized_header_key, "").strip()) < 20
                    if len_check and structure_check:
                        header_matched = True
                        match_method = "in"

                # If matched by either method:
                if header_matched:
                    current_category = category_value
                    is_header = True
                    # print(f"  Found Header: '{header_key}' -> Category: '{current_category}' (Method: {match_method})") # Debug
                    break # Stop checking once the first (most specific due to order) match is found

            if is_header:
                continue

            no, pkt, nama = parse_personnel_line(original_text)
            if no:
                # Name found under a recognised category
                assigned_category = current_category if current_category else NEW_CATEGORY_LABEL
                rows.append([assigned_category, no, pkt, nama, unit_name, ""])
                # print(f"    Extracted Personnel: {no}, {pkt}, {nama} -> Category: {assigned_category}") # Debug
            else:
                # If unable to recognised personnel data and no known category:
                # Reset category so subsequent personnel fall under "NEW CHECK"
                # print(f"  Unrecognized Line (Resetting Category): '{text_for_headers}'") # Debug
                current_category = None

    except FileNotFoundError:
        print(f"Error: Document not found at '{filepath}'")
    except PackageNotFoundError:
        print(f"Error: Could not read '{filepath}'. File might be corrupted or not a valid DOCX.")
    except Exception as e:
        print(f"Error processing document '{filepath}': {e}")
    return rows

# --- Main Execution ---
def main():
    if not os.path.isdir(INPUT_FOLDER):
        print(f"Error: Input folder '{INPUT_FOLDER}' not found.")
        return

    all_personnel_data = []

    print(f"Starting processing in folder: '{INPUT_FOLDER}'")
    for filename in os.listdir(INPUT_FOLDER):
        if filename.lower().endswith(".docx") and not filename.startswith("~"):
            filepath = os.path.join(INPUT_FOLDER, filename)
            unit = os.path.splitext(filename)[0].upper()
            doc_rows = process_document(filepath, unit)
            all_personnel_data.extend(doc_rows)

    if not all_personnel_data:
        print("No personnel data extracted. Exiting.")
        return

    df = pd.DataFrame(all_personnel_data, columns=["CATEGORY", "NO", "PKT", "NAMA", "UNIT", "CATITAN"])

    # Add total count under ('BIL') column
    df.insert(1, "BIL", range(1, len(df) + 1))

    # --- Final Checks & Reporting ---
    new_check_count = len(df[df["CATEGORY"] == NEW_CATEGORY_LABEL])
    if new_check_count > 0:
        print(f"\nWARNING: Found {new_check_count} entries assigned to '{NEW_CATEGORY_LABEL}'. Review required.")

    # Export to Excel as Output_1.xlsx
    try:
        df.to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
        print(f"\n✅ Exported total of {len(df)} unique names to '{OUTPUT_FILE}'")
    except PermissionError:
        print(f"\nError: Permission denied to write to '{OUTPUT_FILE}'. File might be open or directory locked.")
    except Exception as e:
        print(f"\nError writing Excel file '{OUTPUT_FILE}': {e}")

    # For 2nd output: 
    # --- Sort by Category and Add Per-Category 'BIL' ---
    df = df.sort_values(by="CATEGORY", kind='stable').reset_index(drop=True)
    df.insert(1, 'COUNT', df.groupby('CATEGORY').cumcount() + 1)
    # --- End Sorting and Counting ---

        # Export to Excel as Output_sorted.xlsx
    try:
        df.to_excel(OUTPUT_SORTED_FILE, index=False, engine='openpyxl')
        print(f"\n✅ Exported total of {len(df)} unique names sorted by category to '{OUTPUT_SORTED_FILE}'")
        print(f"\nNOTE: Personnel names split across multiple lines/paragraphs may not be fully captured.")
    except PermissionError:
        print(f"\nError: Permission denied to write to '{OUTPUT_SORTED_FILE}'. File might be open or directory locked.")
    except Exception as e:
        print(f"\nError writing Excel file '{OUTPUT_SORTED_FILE}': {e}")

if __name__ == "__main__":
    main()
