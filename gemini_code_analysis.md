# Gemini Code Analysis - Cost Sheet Application

This document outlines the analysis of the Python application for generating project cost sheets, as performed by the Gemini AI assistant.

## Overview

The application is a Streamlit web app designed to streamline the creation and management of project cost sheets. It focuses on ventilation systems (canopies, RECOAIR, SDU, UV-C) and associated components like fire suppression and wall cladding. The app allows users to:

1.  **Create New Projects:** Input project details through a form, which then generates Excel cost sheets and Word quotation documents from templates.
2.  **Revise Existing Projects:** Upload an existing Excel cost sheet, make modifications (add new areas/floors, create new revisions, edit canopies), and download the updated Excel.
3.  **Upload and Generate:** Upload an existing Excel project, and directly generate updated Excel and Word documents, including a final ZIP archive.

## Core File: `app.py`

The primary logic resides in `app.py` (approx. 4417 lines).

### Key Data Structures:

- **`st.session_state.project_data`**: A dictionary built by the "Create New Project" Streamlit form. It holds hierarchical data: Project -> Levels -> Areas -> Canopies, along with their attributes.
- **`project_data` (from `read_excel_file`)**: A dictionary populated by parsing an uploaded Excel file. Its typical structure is:
  ```python
  {
      'sheets': [
          { # Sheet 1 data
              'sheet_name': '...',
              'revision': '...',
              'project_info': {'project_number': '...', 'customer': '...', ...},
              'canopies': [{'reference_number': '...', 'model': '...', ...}, ...],
              'fire_suppression_items': [...],
              'total_price': ...,
              # ... other sheet-specific totals and data
          },
          # ... other sheets
      ],
      'global_fs_k9_total': ...,
      'global_fs_n9_total': ...,
      'global_uv_k9_total': ...,
      # ... other global totals and flags (e.g., 'has_recoair', 'has_sdu')
  }
  ```
  This structure is primarily used for generating Word documents and populating the "JOB TOTAL" Excel sheet.

### External Resources (Templates):

- **Excel:** `resources/Halton Cost Sheet Jan 2025.xlsx`
- **Word (for `docxtpl`):**
  - `resources/Halton Quote Feb 2024 (1).docx` (Main quotation)
  - `resources/Halton RECO Quotation Jan 2025 (2).docx` (RecoAir specific quotation)

### Main Libraries Used:

- **Streamlit:** For the web application interface.
- **Openpyxl:** For reading and writing Excel files (`.xlsx`).
- **Docxtpl:** For generating Word documents from templates using Jinja2-like syntax.
- **Pandas:** Imported, but its direct use in the core generation logic is not prominent.
- **os, shutil, datetime, math, zipfile, re:** Standard Python libraries for file operations, date/time, calculations, ZIP archives, and regular expressions.

## Application Flow:

The `main()` function in `app.py` sets up three Streamlit tabs:

1.  **"Create New Project" (`create_general_info_form()`):**

    - User fills out a detailed form capturing project hierarchy and specifications. Data is stored in `st.session_state.project_data`.
    - Upon submission (e.g., "Generate Cost Sheet" button):
      - `save_to_excel(st.session_state.project_data)`:
        - Loads the Excel template.
        - Iterates through levels/areas/canopies from form data.
        - Uses existing template sheets (CANOPY, FIRE SUPP, EBOX, RECOAIR, SDU) or errors if not enough are available.
        - `write_to_sheet()` populates these sheets with data and formatting (tab colors).
        - `add_dropdowns_to_sheet()` adds data validation lists.
        - (Potentially) `write_job_total()` is called, possibly after an intermediate step to transform `st.session_state.project_data` into the structure `write_job_total` expects or the function adapts.
        - `organize_sheets_by_area()` reorders sheets.
        - Saves as `output.xlsx`.
      - `generate_word_document(project_data_from_excel_extraction)`:
        - This function likely gets its `project_data` by calling `read_excel_file('output.xlsx')` to ensure the data structure is what `write_to_word_doc` expects.
        - `write_to_word_doc()`: Populates the Word template (`Halton Quote Feb 2024 (1).docx`) using the extracted data. Saves as `output.docx`.
      - `write_to_recoair_doc()`: If RECOAIR units are included, this generates a separate RECOAIR quotation.
      - `create_download_zip()`: Bundles `output.xlsx`, `output.docx`, and the RECOAIR doc (if any) into a ZIP file for download.

2.  **"Revise Project" (`create_revision_tab()`):**

    - User uploads an existing project Excel file.
    - `read_excel_file(uploaded_file)` parses it into the `project_data` dictionary structure.
    - UI options appear to:
      - **Add New Floor/Area (`add_new_floor_area()`):** Finds an empty "CANOPY" sheet in the uploaded Excel (or errors if none), renames/titles it, sets revision, adds dropdowns, and saves as a new `_updated.xlsx` file.
      - **Create New Revision (`create_new_revision()`):** Increments the revision letter (e.g., A to B), updates the revision marker in all CANOPY and JOB TOTAL sheets, re-adds dropdowns, and saves the entire workbook into a new versioned folder and filename (e.g., `Project - Num (Revision B)/original_RevB.xlsx`).
      - Other editing functions (e.g., `edit_floor_area_name`, `add_new_canopy`, `edit_canopy`, `update_cladding`) follow a similar pattern of modifying a temporary copy of the uploaded Excel and saving it as a new file.
    - Each operation provides a download link for the modified Excel file.

3.  **"Upload Project" (`create_upload_project_tab()` -> `create_upload_section()`):**
    - User uploads an existing project Excel file.
    - `read_excel_file(uploaded_file)` parses it into `project_data`.
    - A "Generate Documents" button triggers:
      - `write_to_word_doc()`: Generates the main Word quotation directly from the extracted `project_data`.
      - The uploaded Excel is saved to a temporary path. `write_job_total()` is called on this temporary Excel to ensure totals are correct. The temp Excel is saved.
      - `create_download_zip()`: Bundles the (potentially updated) temporary Excel and the generated Word documents (including a potentially customized RECOAIR quotation via `write_to_recoair_doc` logic embedded within `create_download_zip`) into a ZIP file.
    - This tab focuses on direct document generation from an uploaded file rather than re-populating the form for UI editing.

## Key Functions and Logic:

- **`create_general_info_form()`:** Builds the Streamlit form for new project data entry. Populates `st.session_state.project_data`.
- **`save_to_excel(data)`:** Takes data (from form or potentially other sources), loads the Excel template, and populates it by creating/managing sheets for canopies, fire suppression, EBOX (UV-C), RECOAIR, and SDU based on the input data.
- **`write_to_sheet(sheet, data, level_name, area_name, canopies, fs_sheet, is_edge_box)`:** Core function to write detailed project and canopy information to specific cells in an Excel sheet. Differentiates logic for "edge box" type sheets (EBOX, RECOAIR, SDU).
- **`add_dropdowns_to_sheet(wb, sheet, start_row)`:** Adds data validation dropdown lists to canopy sheets, sourcing options from a hidden 'Lists' sheet.
- **`add_recoair_internal_external_dropdowns(sheet)`:** Specific dropdown setup for RECOAIR sheets.
- **`write_job_total(workbook, project_data)`:** Populates the "JOB TOTAL" sheet with summarized project information and consolidated totals from all other relevant sheets and global data. It expects `project_data` in the format returned by `read_excel_file`.
- **`organize_sheets_by_area(workbook)`:** Reorders sheets in the Excel workbook to group related area sheets (CANOPY, FIRE SUPP, EBOX, RECOAIR, SDU) together.
- **`generate_word_document(project_data)`:** Prepares data and calls `write_to_word_doc`. Expects `project_data` from `read_excel_file`.
- **`write_to_word_doc(data, project_data, output_path)`:** Uses `DocxTemplate` to populate `resources/Halton Quote Feb 2024 (1).docx` with project details, canopy information, cladding, MUA calculations, and other specifics.
- **`write_to_recoair_doc(data, project_data, output_path)`:** Generates a RECOAIR-specific quotation using `resources/Halton RECO Quotation Jan 2025 (2).docx` if RECOAIR units are present.
- **`read_excel_file(uploaded_file)`:** Loads an uploaded workbook and calls `extract_sheet_data` for each relevant sheet to parse its contents into a structured `project_data` dictionary.
- **`extract_sheet_data(sheet)`:** The most complex data extraction function. Reads numerous specific cells from a single Excel sheet to gather project info, canopy details (dimensions, prices, model, configuration), fire suppression data (by cross-referencing a "FIRE SUPP" sheet), wall cladding, MUA volumes (with specific calculation logic for 'F' type canopies), and water wash details.
- **`create_download_zip(...)`:** Creates a ZIP archive containing the generated/updated Excel file, the main Word quotation, and the RECOAIR Word quotation (if applicable).
- **Revision/Upload Helper Functions (`create_new_revision`, `add_new_floor_area`, etc.):** These functions typically work on a temporary copy of an uploaded Excel file, make specific modifications using `openpyxl`, and save the result as a new file.

## Points of Note:

- **Data Model Transformation:** A key aspect is the transformation of data from the Streamlit form's structure (`st.session_state.project_data`) to the structure expected by Word generation and totalization functions (the output of `read_excel_file`). This likely happens by saving the form data to Excel and then immediately re-reading it.
- **Template Dependency:** The system is tightly coupled to the structure of the Excel and Word templates. Changes in template cell locations or placeholder names would require code modifications.
- **Error Handling:** `try-except` blocks are used throughout to manage potential issues during file operations and data processing, with errors often reported to the Streamlit UI.
- **MUA Calculation:** Specific logic exists in `extract_sheet_data` to calculate Make-Up Air (MUA) volume for 'F' type canopies, comparing an 85% calculation against a 'MAX SUPPLY' value from the sheet.
- **Fire Suppression Cost Apportionment:** When extracting fire suppression data, the total installation cost from the "FIRE SUPP" sheet is apportioned among the canopies listed on that sheet.

This analysis provides a comprehensive understanding of the application's architecture and execution flow.
