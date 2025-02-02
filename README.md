# vba-data-transformation

These VBA projects automate the transformation of raw purchase data generated from ERP systems (such as Microsoft, Oracle, SAP Hana) into a standardized government or ingestion template with a single button click.

**Key Features**

- *Automated Data Processing*: Reads input from an ERP-generated Excel file and formats it to match the required government or default ingestion template.
- *Sorting & Data Cleaning*: Organizes data by key columns and ensures structured output.
- *Template Application*: Maps essential fields, including invoice details, customer information, and GSTIN validation.
- *File Management*: Opens, processes, and saves output files without manual intervention.
- *Password-Protected Sheets*: Unlocks and modifies pre-defined templates securely.
- *Data Validations*: Apply custom logic for validations (i.e. specific chartacter limit, data type, GSTIN Verification)

**Customization**: This script can be easily modified to accommodate different formats, additional fields, or validation logic based on business requirements.

**🚨Note on Sample**

Sorry, I can't add the .xlsx files as they contain confidential GST financial data. However, you can refer to the syntax and logic to customize it according to your needs. The core approach remains the same.

**⚠  Common Mistakes to avoid**
- *File Naming & Extensions* – Ensure the correct Excel file names and extensions (.xlsx, .xlsm). Avoid .xml or incorrect formats.
- *File Paths* – Double-check the full file path before running the macro to prevent file not found errors.
- *Understanding Nested Logic* – Ensure correct usage of If, For, and other VBA functions to avoid unexpected results.
- *Sheet References* – Make sure you're copying and pasting from the correct sheets within the workbooks.
- *Unprotected Sheets* – If modifying a protected sheet, use the correct password to avoid runtime errors.
