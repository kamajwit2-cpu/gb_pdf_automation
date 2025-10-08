# PDF Tabular Data Extractor

This project extracts structured tabular data from PDF files, specifically designed for jewelry manufacturing documents.

## Features

- **Comprehensive Data Extraction**: Extracts SKU codes, metal specifications, diamond quality, dates, quantities, and more
- **Multi-page Support**: Handles both single and multiple page PDFs
- **Pattern Recognition**: Uses advanced regex patterns to identify jewelry SKU codes and specifications
- **Excel Output**: Saves extracted data in structured Excel format
- **Batch Processing**: Processes all PDFs in a folder automatically

## Setup

### 1. Virtual Environment
```bash
# Create virtual environment
python -m venv pdf_extraction_env

# Activate virtual environment (Windows)
pdf_extraction_env\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Directory Structure
```
Python GB PDF Automation/
├── Mail Download/          # Place your PDF files here
├── Python Output/          # Extracted data will be saved here
├── pdf_tabular_extractor.py
├── requirements.txt
└── README.md
```

## Usage

### Process All PDFs
```bash
python pdf_tabular_extractor.py
```

This will:
- Process all PDF files in the `Mail Download` folder
- Extract structured data from each PDF
- Save individual Excel files for each PDF
- Create a combined Excel file with all extracted data

### Process Individual PDF
```python
from pdf_tabular_extractor import PDFTabularExtractor

# Initialize extractor
extractor = PDFTabularExtractor()

# Process a single PDF
success = extractor.process_single_pdf("your_file.pdf")
```

## Extracted Data Fields

The extractor captures the following information:

### Header Information
- **Vendor**: Purchase from vendor name
- **PO Number**: Purchase order number
- **PO Date**: Purchase order date
- **Vendor PO**: Vendor's purchase order number
- **Customer Name**: Customer name

### Product Specifications
- **Kama SKU**: Product SKU code
- **Metal KT**: Metal karat (14KT, 18KT, PLAT, etc.)
- **Metal Color**: Metal color (Yellow, White, Rose, etc.)
- **Size**: Product size
- **Dia Quality**: Diamond quality (F-VS2, G+/VS, etc.)
- **ItemType**: Item type
- **Engraving**: Engraving text
- **Desc**: Product description
- **Category**: Product category

### Weights and Tolerances
- **Metal Tol.**: Metal weight tolerance range
- **Dia Tol.**: Diamond weight tolerance range

### Order Information
- **Ship Date**: Required ship date
- **Qty**: Quantity
- **Line No**: Line number (101, 102, etc.)

## Output Files

### Individual Files
For each PDF processed, the following files are created:
- `{filename}_text.txt`: Raw extracted text
- `{filename}_tabular_data.xlsx`: Structured data in Excel format

### Combined File
- `combined_tabular_data_{timestamp}.xlsx`: All extracted data combined into one Excel file

## Technical Details

### SKU Pattern Recognition
The extractor recognizes over 100 different SKU patterns including:
- A-series: A11, A07, A08, A32, AJR
- B-series: B89, B09, B74, B19, B04, B99, B29, BNG, B14, B17
- L-series: LGD, LGR, LGT, LTB, LGB, LGY, LGO, LGX, etc.
- E-series: E11, E36, E50, EAG, E31, E19, E20, E74, E89, E01, E44
- And many more...

### Text Processing Logic
1. **Header Extraction**: Identifies vendor, PO, and customer information
2. **Block Processing**: Splits text into blocks based on line numbers (101, 102, etc.)
3. **Pattern Matching**: Uses regex to extract specific data fields
4. **Data Validation**: Cleans and validates extracted data
5. **Excel Generation**: Creates structured Excel output

## Error Handling

The extractor includes comprehensive error handling:
- Invalid PDF files are skipped with error messages
- Malformed data is handled gracefully
- Missing fields are filled with appropriate defaults
- Encoding issues are resolved automatically

## Dependencies

- `pdfplumber`: PDF text extraction
- `pandas`: Data manipulation and Excel export
- `openpyxl`: Excel file handling
- `numpy`: Numerical operations
- `Pillow`: Image processing (for PDF handling)

## Troubleshooting

### Common Issues

1. **No data extracted**: Check if PDF contains text (not just images)
2. **Encoding errors**: Ensure PDF files are not corrupted
3. **Missing dependencies**: Run `pip install -r requirements.txt`
4. **Permission errors**: Ensure write access to output folder

### Performance Tips

- Process PDFs in batches for large collections
- Use SSD storage for faster file I/O
- Close other applications to free up memory for large PDFs

## Support

For issues or questions, check the original `pdf 9.py` file for reference logic or contact the development team.

