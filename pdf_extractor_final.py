"""
Final PDF Extractor - Based on Actual PDF Structure Analysis
===========================================================

This script extracts data from PDF files based on the actual structure:
- Common header on every page with PO info, vendor info, customer info
- Line items with sequential numbers (101, 102, 103...)
- Kama SKU is the 4th field in each line item
- Handles both single and multiple line items
- Combines all PDFs into a single Excel file
"""

import datetime
from datetime import datetime
import pandas as pd
import pdfplumber
import os
import re
from typing import Dict, List, Optional, Any


class FinalPDFExtractor:
    """Final PDF extractor based on actual PDF structure analysis."""
    
    def __init__(self, input_folder: str = None, output_folder: str = "Python Output"):
        """
        Initialize the final PDF extractor.
        
        Args:
            input_folder: Path to folder containing PDF files (None for production mode)
            output_folder: Path to folder for saving extracted data
        """
        # Environment detection
        self.is_production = os.getenv('ENV') == 'production'
        
        if input_folder is None:
            if self.is_production:
                # Production mode: will use uploaded files
                self.input_folder = None
            else:
                # Development mode: use Mail Download folder
                self.input_folder = "Mail Download"
        else:
            self.input_folder = input_folder
            
        self.output_folder = output_folder
        self.ensure_output_folder()
    
    def ensure_output_folder(self):
        """Ensure the output folder exists."""
        os.makedirs(self.output_folder, exist_ok=True)
    
    def list_pdf_files(self) -> List[str]:
        """List all PDF files in the input folder."""
        if self.input_folder is None:
            # Production mode: no folder-based processing
            return []
            
        if not os.path.exists(self.input_folder):
            print(f"Input folder '{self.input_folder}' does not exist!")
            return []
        
        all_files = os.listdir(self.input_folder)
        pdf_files = [file for file in all_files if file.lower().endswith('.pdf')]
        return pdf_files
    
    def extract_header_info(self, text: str) -> Dict[str, str]:
        """
        Extract header information from the common header section.
        
        Args:
            text: Full text from PDF
            
        Returns:
            Dictionary with header information
        """
        header_info = {
            "Vendor": "",
            "PO Number": "",
            "PO Date": "",
            "Vendor PO": "",
            "Customer Name": ""
        }
        
        # Extract Vendor (from Purchase From)
        purchase_match = re.search(r"Purchase From\s*:\s*([^\s]+)", text)
        if purchase_match:
            header_info["Vendor"] = purchase_match.group(1)
        
        # Extract PO Number
        po_match = re.search(r"PO #\s*:\s*(\d+)", text)
        if po_match:
            header_info["PO Number"] = po_match.group(1)
        
        # VPO is now split into PO Date and Vendor PO - no longer extracting combined VPO
        
        # Extract Customer Name
        cust_match = re.search(r"Cust Name\s+([^\n]+)", text)
        if cust_match:
            header_info["Customer Name"] = cust_match.group(1).strip()
        
        # Extract PO Date
        po_date_match = re.search(r"Date\s*:\s*(\w+/\d+/\d+)", text)
        if po_date_match:
            try:
                header_info["PO Date"] = datetime.strptime(po_date_match.group(1), "%b/%d/%Y").strftime("%d-%m-%Y")
            except ValueError:
                header_info["PO Date"] = po_date_match.group(1)
        
        # Extract Vendor PO
        vendor_po_match = re.search(r"Vendor PO #\s+([^\n]+)", text)
        if vendor_po_match:
            header_info["Vendor PO"] = vendor_po_match.group(1).strip()
        
        return header_info
    
    def parse_line_item(self, full_context: str) -> Optional[Dict[str, str]]:
        """
        Parse a single line item to extract all relevant data.
        
        Args:
            full_context: Full context text including line and following lines
            
        Returns:
            Dictionary with line item data or None if invalid
        """
        # Get the first line from the full context
        first_line = full_context.split('\n')[0]
        
        # Check if line starts with a line number (101-999, but exclude common header numbers)
        # Handle both formats: "101 2206492/101" and "101 E2831893/101"
        line_match = re.match(r"^(\d{3})\s*(.+)", first_line)
        if not line_match:
            return None
        
        line_number = int(line_match.group(1))
        
        # Skip common header numbers that appear in addresses
        if line_number < 101 or line_number > 999:
            return None
        
        rest_of_line = line_match.group(2)
        
        # Split the rest of the line by spaces
        parts = rest_of_line.split()
        
        if len(parts) < 3:
            return None
        
        # Extract fields based on position
        # Handle different formats:
        # Format 1: "101 2206492/101 TXE01899-WGGD LGD-TXE01899-GW4 14KT W"
        # Format 2: "101 E2831893/101LGR-RVR07619-WG4 LGR-RVR07619-WG4 14KT W"
        
        order_number = parts[0] if len(parts) > 0 else ""
        
        # Check if the second part contains both order number and SKU (no space)
        # Handle "stuck together" pattern: "E2831893/101LGR-RVR07619-WG4 LGR-RVR07619-WG4"
        if len(parts) > 2 and "/" in parts[1] and not parts[1].startswith("LGD") and not parts[1].startswith("LG") and not parts[1].startswith("LGR"):
            # This is the combined format: "E2831893/101LGR-RVR07619-WG4"
            combined = parts[1]
            # Find where the SKU starts (usually after the last slash and number)
            match = re.search(r'(.+?/\d+)(.+)', combined)
            if match:
                order_number = match.group(1)  # "E2831893/101"
                vendor_style = match.group(2)  # "LGR-RVR07619-WG4" (the SKU part stuck in parts[1])
                kama_sku = parts[2]  # Use parts[2] as the actual SKU (clean version)
            else:
                vendor_style = parts[1]
                kama_sku = parts[2] if len(parts) > 2 else ""
        else:
            # Check if parts[1] looks like a SKU (contains letters and numbers, not just metal info)
            if len(parts) > 1 and re.search(r'[A-Z]', parts[1]) and not re.match(r'^\d{2}KT$', parts[1]):
                # parts[1] is the SKU, parts[2] is metal info
                vendor_style = parts[1]
                kama_sku = parts[1]  # SKU is in parts[1]
            else:
                # Normal format: use parts[1] as vendor style and parts[2] as SKU
                vendor_style = parts[1] if len(parts) > 1 else ""
                kama_sku = parts[2] if len(parts) > 2 else ""
        
        # Validate that we have a Kama SKU (any pattern is acceptable)
        if not kama_sku or len(kama_sku) < 3:
            return None
        
        # Additional validation: Skip lines that don't have proper product information
        # Lines should have Metal KT (like 14KT, 18KT, etc.) to be valid product lines
        if not any(metal_kt in rest_of_line for metal_kt in ["KT", "PLAT", "SILV", "STER"]):
            return None
        
        # Find metal specifications (KT, PLAT, etc.)
        metal_kt = ""
        metal_color = ""
        size = ""
        
        # Look for metal patterns in all parts
        for i, part in enumerate(parts):
            if re.match(r"\d{2}KT|PLAT|SILV|STER", part):
                metal_kt = part
                if i + 1 < len(parts):
                    metal_color = parts[i + 1]
                break
        
        
        # Find size - look for decimal numbers in the line (like 1.0000, 7.25, etc.)
        for part in parts:
            size_match = re.match(r"^(\d+(?:\.\d+)?)$", part)
            if size_match:
                size = f"{float(size_match.group(1)):.2f}"
                break
        
        # Find diamond quality (pattern like F-VS2, E-VS1, D/VVS EX, etc.)
        dia_quality = ""
        for part in parts:
            if re.match(r"[A-Z][+-]?[-/][A-Z0-9]+", part):
                dia_quality = part
                break
        
        # Extract product category from SKU patterns
        category = ""
        sku_upper = kama_sku.upper()
        
        # Check for specific patterns in SKU
        if "BAND" in sku_upper or sku_upper.startswith("LGB-"):
            category = "BAND"
        elif "BRAC" in sku_upper or "BRACELET" in sku_upper:
            category = "BRACELET"
        elif "NECK" in sku_upper or "NECKLACE" in sku_upper:
            category = "NECKLACE"
        elif "PEND" in sku_upper or "PENDANT" in sku_upper or "-RVP" in sku_upper:
            category = "PENDANT"
        elif "RING" in sku_upper or "-WR" in sku_upper or "-RVR" in sku_upper or "-TXR" in sku_upper:
            category = "RING"
        elif "EARR" in sku_upper or "EARRING" in sku_upper or "-TXE" in sku_upper or "-RVE" in sku_upper:
            category = "EARRING"
        elif "CHAIN" in sku_upper:
            category = "CHAIN"
        elif "SET" in sku_upper:
            category = "SET"
        
        # Check for metal pattern indicators (like 18KP = 18K Pendant, 14KWR = 14K Ring)
        if not category:  # Only if category not found from SKU
            # Look for patterns like 18KP, 14KWR, etc. in the line
            metal_pattern = re.search(r"(\d{2})K\s*([A-Z]+)", rest_of_line)
            if metal_pattern:
                # Don't overwrite metal_kt - it's already correctly extracted above
                suffix = metal_pattern.group(2)
                
                if "P" in suffix or "PEND" in suffix:
                    category = "PENDANT"
                elif "R" in suffix or "RING" in suffix:
                    category = "RING"
                elif "B" in suffix or "BAND" in suffix:
                    category = "BAND"
                elif "N" in suffix or "NECK" in suffix:
                    category = "NECKLACE"
                elif "E" in suffix or "EARR" in suffix:
                    category = "EARRING"
                elif "C" in suffix or "CHAIN" in suffix:
                    category = "CHAIN"
        
        # Find date and quantity
        ship_date = ""
        qty = ""
        
        for i, part in enumerate(parts):
            if re.match(r"\w{3}/\d{2}/\d{4}", part):
                try:
                    ship_date = datetime.strptime(part, "%b/%d/%Y").strftime("%d-%m-%Y")
                except ValueError:
                    ship_date = part
                
                # Quantity is usually the next part
                if i + 1 < len(parts):
                    qty = parts[i + 1]
                break
        
        # Extract additional fields
        item_type = ""
        engraving = ""
        desc = ""
        metal_tol = ""
        dia_tol = ""
        req_date = ""
        
        # Find ItemType
        itemtype_match = re.search(r"ItemType\s+(\S+)", full_context)
        if itemtype_match:
            item_type = itemtype_match.group(1)
        
        # Find Engraving (using your original logic)
        engraving_match = re.search(r"Engraving Text:\s*(.*)", full_context)
        if engraving_match:
            engraving = engraving_match.group(1).strip()
        
        # Find Description
        desc_match = re.search(r"Desc\.\s*(.*?)\s*GM Min", full_context)
        if desc_match:
            desc = desc_match.group(1).strip()
        
        # If category not found from SKU patterns, try to extract from Description
        if not category and desc:
            desc_upper = desc.upper()
            if "EARRING" in desc_upper:
                category = "EARRING"
            elif "WEDDING BAND" in desc_upper or "MACHINE BAND" in desc_upper or "BAND" in desc_upper:
                category = "BAND"
            elif "BRACELET" in desc_upper or "BRACLET" in desc_upper or "BCG" in desc_upper:
                category = "BRACELET"
            elif "PENDANT" in desc_upper or " PD" in desc_upper or "18KP" in desc_upper:
                category = "PENDANT"
            elif "NECKLACE" in desc_upper or "NECKALCE" in desc_upper or "NACKLACE" in desc_upper:
                category = "NECKLACE"
            elif "RING" in desc_upper or " R " in desc_upper:
                category = "RING"
            elif "CHAIN" in desc_upper:
                category = "CHAIN"
            elif "SET" in desc_upper:
                category = "SET"
        
        # Find Metal Tolerance
        metal_tol_match = re.search(r"GM\s+Min Wt\s+([\d.]+)\s+Max Wt\s+([\d.]+)", full_context)
        if metal_tol_match:
            metal_tol = f"{metal_tol_match.group(1)}-{metal_tol_match.group(2)}"
        
        # Find Diamond Tolerance
        dia_tol_match = re.search(r"Diam\s+Min Wt\s+([\d.]+)\s+Max Wt\s+([\d.]+)", full_context)
        if dia_tol_match:
            dia_tol = f"{dia_tol_match.group(1)}-{dia_tol_match.group(2)}"
        
        # Req Date is same as Ship Date for now
        req_date = ship_date
        
        return {
            "Line No": str(line_number),
            "Order Number": order_number,
            "Vendor Style": vendor_style,
            "Kama SKU": kama_sku,
            "Metal KT": metal_kt,
            "Metal Color": metal_color,
            "Size": size,
            "Ship Date": ship_date,
            "Qty": qty,
            "Dia Quality": dia_quality,
            "ItemType": item_type,
            "Engraving": engraving,
            "Desc": desc,
            "Category": category,
            "Metal Tol.": metal_tol,
            "Dia Tol.": dia_tol,
            "Req Date": req_date
        }
    
    def extract_line_items(self, text: str) -> List[Dict[str, str]]:
        """
        Extract all line items from the text.
        
        Args:
            text: Full text from PDF
            
        Returns:
            List of line item dictionaries
        """
        line_items = []
        lines = text.splitlines()
        
        for i, line in enumerate(lines):
            # Look for lines starting with line numbers (101, 102, etc.)
            if re.match(r"^\d{3}\s+", line):
                # Get the full context for this line item (current line + next few lines)
                full_context = line
                for j in range(i + 1, min(i + 10, len(lines))):
                    if re.match(r"^\d{3}\s+", lines[j]):
                        break
                    full_context += "\n" + lines[j]
                
                line_data = self.parse_line_item(full_context)
                if line_data:
                    line_items.append(line_data)
                # Line was rejected (no debug output needed)
        
        return line_items
    
    def process_single_pdf(self, pdf_path: str) -> List[Dict[str, Any]]:
        """
        Process a single PDF file (for production mode with uploaded files).
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            List of dictionaries containing extracted data
        """
        return self.extract_data_from_pdf(pdf_path)
    
    def extract_data_from_pdf(self, pdf_path: str) -> List[Dict[str, Any]]:
        """
        Extract all data from a single PDF file.
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            List of dictionaries containing extracted data
        """
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text(x_tolerance=1, y_tolerance=1) + "\n"
            
            # Save extracted text to file for debugging
            pdf_filename = os.path.basename(pdf_path)
            text_filename = pdf_filename.replace('.pdf', '_text.txt')
            text_output_path = os.path.join(self.output_folder, text_filename)
            
            with open(text_output_path, 'w', encoding='utf-8') as f:
                f.write(f"EXTRACTED TEXT FROM: {pdf_filename}\n")
                f.write("=" * 80 + "\n\n")
                f.write(text)
                f.write("\n" + "=" * 80 + "\n")
                f.write("END OF EXTRACTED TEXT\n")
            
            # Extract header information
            header_info = self.extract_header_info(text)
            
            # Extract line items
            line_items = self.extract_line_items(text)
            
            # Combine header info with each line item
            combined_data = []
            for line_item in line_items:
                combined_item = {**header_info, **line_item}
                combined_data.append(combined_item)
            
            return combined_data
            
        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")
            return []
    
    def process_all_pdfs(self) -> None:
        """Process all PDF files and save combined data to Excel."""
        pdf_files = self.list_pdf_files()
        
        if not pdf_files:
            print("No PDF files found in the input folder!")
            return
        
        print(f"Found {len(pdf_files)} PDF files to process...")
        
        all_data = []
        successful_count = 0
        
        for pdf_file in pdf_files:
            pdf_path = os.path.join(self.input_folder, pdf_file)
            print(f"Processing: {pdf_file}")
            
            extracted_data = self.extract_data_from_pdf(pdf_path)
            
            if extracted_data:
                all_data.extend(extracted_data)
                successful_count += 1
                print(f"  Extracted {len(extracted_data)} line items")
                print(f"  Text saved to: {pdf_file.replace('.pdf', '_text.txt')}")
            else:
                print(f"  No data extracted")
        
        # Save combined data
        if all_data:
            current_datetime = datetime.now().strftime("%d.%m.%Y_%H.%M")
            output_excel_file_path = os.path.join(
                self.output_folder, 
                f"final_combined_data_{current_datetime}.xlsx"
            )
            
            try:
                df = pd.DataFrame(all_data)
                
                # Reorder columns as requested
                column_order = [
                    "Vendor", "PO Number", "PO Date", "Vendor PO", "Customer Name", "Kama SKU", 
                    "Metal KT", "Metal Color", "Size", "Ship Date", "Qty", 
                    "Dia Quality", "ItemType", "Engraving", "Desc", "Category", 
                    "Line No", "Metal Tol.", "Dia Tol.", "Req Date"
                ]
                
                # Only include columns that exist in the dataframe
                existing_columns = [col for col in column_order if col in df.columns]
                df = df[existing_columns]
                
                df.to_excel(output_excel_file_path, index=False)
                
                print(f"\nSuccessfully processed {successful_count}/{len(pdf_files)} PDF files")
                print(f"Total line items extracted: {len(all_data)}")
                print(f"Combined data saved to: final_combined_data_{current_datetime}.xlsx")
                
                # Show sample data
                print(f"\nSample data:")
                print(df[['Vendor', 'PO Number', 'PO Date', 'Vendor PO', 'Customer Name', 'Kama SKU', 'Metal KT', 'Metal Color', 'Size']].head(10).to_string())
                
            except Exception as e:
                print(f"Error saving Excel file: {e}")
        else:
            print("No data was extracted from any PDF files!")


def main():
    """Main function to run the final PDF extractor."""
    print("Final PDF Extractor - Based on Actual PDF Structure")
    print("=" * 60)
    
    # Initialize extractor
    extractor = FinalPDFExtractor()
    
    # Process all PDFs
    extractor.process_all_pdfs()


if __name__ == "__main__":
    main()
