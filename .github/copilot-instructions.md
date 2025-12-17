# SYSMEX Workplace - SAP Excel Processor

## Project Purpose
This is a specialized data processing tool for SAP inventory exports. It transforms messy SAP Excel files into clean, standardized inventory reports with automatic filtering, sorting, and Excel table formatting.

## Core Architecture

### Single-Purpose Script Design
- **Main processor**: [sap_processor.py](../sap_processor.py) - standalone CLI tool, no classes/modules
- **Test suite**: [test_sap_processor.py](../test_sap_processor.py) - unittest framework with full integration tests
- Three functions handle the pipeline: `normalize_columns()` → `find_best_sheet()` → `process_file()`

### Data Flow Pattern
```
Input: sklady_porovnani/input/*.xlsx (raw SAP exports)
  ↓
Process: Auto-detect sheet → Normalize → Filter → Transform → Sort → Format
  ↓
Output: sklady_porovnani/output/*.xlsx (standardized tables with "tbl_SAP")
```

## Critical Conventions

### Column Handling (Case-Insensitive)
The script expects SAP exports with these columns (any case):
- **Required**: Material, Material description, Batch, Total Quantity, Storage location
- **Optional**: Plant (gets dropped)
- See `REQUIRED_COLS` and `OPTIONAL_COLS` dictionaries in [sap_processor.py](../sap_processor.py#L8-L17)

### Fixed Transformation Pipeline (Order Matters!)
1. Filter: Keep only Storage location = F010 or F070
2. Drop: Remove Plant and Storage location columns
3. Reorder: Material → Material description → Batch → Total Quantity
4. Rename: "Material description" → "Název", "Total Quantity" → "Mnozstvi_SAP"
5. Sort: Descending by Mnozstvi_SAP (numeric, NaN→0)
6. **Delete first row** after sorting (critical business logic!)
7. Trim to exactly 4 columns

### Output Format Requirements
- Excel file with sheet named "SAP"
- Data wrapped in Excel Table named "tbl_SAP" (range A1:D{rows})
- Table style: TableStyleMedium9 (blue striped)
- Uses openpyxl for table creation after pandas writes data

## Developer Workflows

### Running the Processor
```bash
# Process all files in sklady_porovnani/input/
python sap_processor.py --mode all

# Process single file
python sap_processor.py --mode single --file_path path/to/file.xlsx
```

### Testing
```bash
# Run full integration test (creates temp files, validates entire pipeline)
python -m unittest test_sap_processor.py

# The test validates: column normalization, filtering, sorting, row deletion,
# table creation, and exact output format
```

### Dependencies
All managed in [requirements.txt](../requirements.txt):
- `pandas` - Data manipulation
- `openpyxl` - Excel table formatting (pandas engine for .xlsx)
- `argparse` - CLI interface (stdlib but listed)

## Language & Comments
- All code comments, print statements, and error messages are in **Czech**
- Variable names follow mix: English (df, mapping) and Czech (Mnozstvi_SAP, Název)
- Keep this bilingual pattern when modifying code

## Common Pitfalls
- **Don't skip the "delete first row after sort" step** - this is intentional business logic, not a bug
- **Sheet auto-detection** uses largest data area, not first sheet or sheet names
- **Storage location filtering** is hardcoded to F010/F070 - these are specific warehouse codes
- **Total Quantity** must handle non-numeric values (converts to 0) before sorting
- Output directory `sklady_porovnani/output/` is auto-created if missing

## Testing Strategy
- Uses temporary files (test_input_temp.xlsx, test_output_temp.xlsx) created/destroyed per test
- Validates both data transformations AND Excel formatting (pandas for data, openpyxl for table structure)
- Test data includes edge cases: case-insensitive columns, non-numeric quantities, invalid storage locations
