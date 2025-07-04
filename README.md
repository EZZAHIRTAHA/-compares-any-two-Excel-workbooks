# Excel Diff Tool 📊

A powerful Python utility for comparing two Excel workbooks and identifying differences between them. Perfect for data analysts, accountants, and anyone who needs to track changes between Excel files.

## ✨ Features

- **Smart Column Alignment**: Automatically handles Excel files with different column structures
- **Comprehensive Comparison**: Identifies added, deleted, and modified rows
- **Flexible Key Columns**: Support for single or multiple columns as unique identifiers
- **Rich Output**: Generates detailed Excel report with separate sheets for each type of change
- **Robust Error Handling**: Gracefully handles edge cases and data inconsistencies

## 🚀 Quick Start

### Installation

```bash
pip install pandas openpyxl
```

### Basic Usage

```bash
python excel_diff.py file1.xlsx file2.xlsx --out differences.xlsx
```

### Advanced Usage

```bash
# Compare specific sheet with custom key columns
python excel_diff.py \
  old_data.xlsx \
  new_data.xlsx \
  --sheet "Sheet1" \
  --key "ID" "Name" \
  --out detailed_diff.xlsx
```

## 📋 Command Line Arguments

| Argument | Description | Example |
|----------|-------------|---------|
| `file_a` | Baseline/older Excel file | `baseline.xlsx` |
| `file_b` | New Excel file to compare | `updated.xlsx` |
| `--sheet` | Sheet name or index (default: first sheet) | `"Data"` or `0` |
| `--key` | Column(s) for unique row identification | `"ID"` or `"Name" "Email"` |
| `--out` | Output Excel file path | `comparison_result.xlsx` |

## 📊 Output Structure

The tool generates an Excel file with three organized sheets:

### 🟢 Added_rows
Rows that exist only in the newer file

### 🔴 Deleted_rows
Rows that exist only in the older file

### 🟡 Modified_cells
Side-by-side comparison of cells that have changed, showing both old and new values

## 🛠️ Technical Details

### How It Works

1. **Data Loading**: Reads specified sheets from both Excel files
2. **Index Setting**: Uses key columns to create unique row identifiers
3. **Column Alignment**: Automatically aligns DataFrames with different column structures
4. **Comparison Logic**: 
   - Identifies rows unique to each file
   - Compares common rows cell-by-cell
   - Tracks all modifications with detailed change tracking

### Key Improvements

- **Column Mismatch Resolution**: Handles files with different column structures
- **Data Type Normalization**: Converts all data to strings for consistent comparison
- **Memory Efficient**: Processes only common rows for modification detection
- **Flexible Indexing**: Works with or without key columns

## 📝 Example Scenarios

### Financial Data Comparison
```bash
python excel_diff.py \
  "Q1_Financial_Report.xlsx" \
  "Q2_Financial_Report.xlsx" \
  --key "Account_ID" \
  --out "Financial_Changes.xlsx"
```

### Inventory Management
```bash
python excel_diff.py \
  "Previous_Inventory.xlsx" \
  "Current_Inventory.xlsx" \
  --key "SKU" "Location" \
  --sheet "Inventory" \
  --out "Inventory_Changes.xlsx"
```

### Customer Database Updates
```bash
python excel_diff.py \
  "Customer_DB_Old.xlsx" \
  "Customer_DB_New.xlsx" \
  --key "Customer_ID" \
  --out "Customer_Updates.xlsx"
```

## 🔧 Troubleshooting

### Common Issues

**"Key column not found"**
- Ensure the specified key columns exist in both files
- Check for typos in column names
- Verify you're comparing the correct sheets

**"Cannot compare DataFrames"**
- The tool automatically handles column mismatches
- If issues persist, check for special characters in column names

**Memory Issues with Large Files**
- Consider processing files in chunks for very large datasets
- Ensure sufficient system memory

## 🤝 Contributing

Feel free to submit issues, feature requests, or pull requests. This tool is designed to be extensible and user-friendly.

## 📄 License

Open source - feel free to use and modify as needed.

---

*Made with ❤️ for data professionals who need reliable Excel comparison tools.*