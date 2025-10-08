import datetime
from datetime import datetime
import pandas as pd
import pdfplumber
import os
import re


def check_101_102_same_line(text):
    """Check if '101' and '102' are in the same line of text."""
    # Split the text into lines
    lines = text.splitlines()

    # Iterate through each line and check for the presence of both '101' and '102'
    for line in lines:
        if "101" in line and "102" in line:
            return True  # Found both in the same line
    return False  # No line contains both


def list_pdf_files(folder_path):
    """List all PDF files in the specified folder."""
    all_files = os.listdir(folder_path)
    pdf_files = [file for file in all_files if file.lower().endswith('.pdf')]
    return pdf_files


def process_single_pdf(pdf_file, folder_path, output_folder):
    """Process a single PDF file by extracting structured data and saving it to text and Excel files."""
    pdf_path = os.path.join(folder_path, pdf_file)
    #print(f"Processing file: {pdf_path}")

    # Extract structured data from the PDF and save text immediately
    output_text_file_name = f"{os.path.splitext(pdf_file)[0]}_text.txt"
    output_text_file_path = os.path.join(output_folder, output_text_file_name)

    extracted_data = extract_data_from_pdf(pdf_path, output_text_file_path)

    if extracted_data:
        # Define the output Excel file path
        output_excel_file_name = f"{os.path.splitext(pdf_file)[0]}_data.xlsx"
        output_excel_file_path = os.path.join(output_folder, output_excel_file_name)

        # Convert the extracted data to a pandas DataFrame and save it as an Excel file
        try:
            df = pd.DataFrame([extracted_data])  # Wrap the data in a list to create a single row DataFrame
            df.to_excel(output_excel_file_path, index=False)
            #print(f"Saved structured data to {output_excel_file_path}")
        except Exception as e:
            print(f"Error saving Excel file {output_excel_file_path}: {e}")


def extract_data_from_pdf(pdf_path, output_text_file_path):
    """Extract structured data from a single PDF file and save text data immediately."""
    # try:
    #     with pdfplumber.open(pdf_path) as pdf:
    #         text = ""
    #         for i, page in enumerate(pdf.pages):
    #             #print(f"Extracting text from Page {i + 1}...")
    #             page_text = page.extract_text()
    #             text += page_text + "\n"
    #
    #         # Immediately save the extracted text to a file
    #         with open(output_text_file_path, 'w', encoding='utf-8') as file:
    #             file.write(text)
    #         #print(f"Saved extracted text to {output_text_file_path}")
    #
    #     # Extract structured data from the text
    #     data = extract_data_from_text(text)
    #     return data
    #
    # except Exception as e:
    #     print(f"Error processing {pdf_path}: {e}")
    #     return None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for i, page in enumerate(pdf.pages):
                # Slightly adjusted tolerance to avoid duplicate/misparsed numbers
                page_text = page.extract_text(x_tolerance=1, y_tolerance=1)
                text += page_text + "\n"

            # Immediately save the extracted text to a file
            with open(output_text_file_path, 'w', encoding='utf-8') as file:
                file.write(text)

        # Continue with your original regex and parsing logic
        data = extract_data_from_text(text)
        return data

    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        return None

def correct_double_decimal_number(val: str) -> str:
    # Match patterns like: 99.5.500, 77.2.255, etc.
    parts = val.split(".")
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        left_digit = parts[0][-1]  # last digit of first part (e.g., '9' from '99')
        right_digits = (parts[2] + "00")[:2]  # first 2 digits of third part, padded if needed
        return f"{left_digit}.{right_digits}"
    return val  # return original if it doesn't match the broken pattern

def process_text(text, data, line_no):
    # Define the patterns we need to search for
    patterns = ["A11", "A07", "A08", "A32", "AJR",
                "B89", "B09", "B74", "B19", "B04", "B99", "B29", "BNG", "B14", "B17",
                "LGD", "L0GD", "L4GD", "LGR", "LGT", "LTB", "LGB", "LGY", "LGO", "LGX", "LG0", "LG-", "LGA", "L6GA",
                "LGW", "LGD-", "LRV", "L\)GD", "LAJ", "LTX", "LGK", "LCR", "LGS", "LGZ", "LGJ", "LGU", "LEN", "LG4",
                "LGC", "LGA", "LGV", "LDT", "LTR", "LGT", "LG3", "LJR", "LR0", "LVR", "LER", "LCW", "LR5", "LGR", "LJE",
                "LRS", "LGEN004W", "LGE",
                "E11", "E36", "E50", "EAG", "E31", "E19", "E20", "E74", "E89", "E01", "E44",
                "F11", "F30", "F34",
                "M54",
                "B92", "G35",
                "G08", "G17", "G11", "G07",
                "D08", "D11", "D96",
                "N25",
                "R09", "R03",  "R10", "R12", "R42", "R53", "R11", "R19", "R51", "R60", "R80", "R04", "R76", "R17", "R77",
                "R00", "RVN", "R84",  "R41", "R98", "R75", "R05", "R92", "R01",
                "VLR", "VLW",
                "P95", "P77", "P16", "P09", "P24", "P13", "P11", "P19", "P12",
                "T12", "T32", "T29", "T14", "T17", "T11", "T08", "TC8", "TC7", "TC6", "T13", "TC9",
                "W35", "W10", "W11", "W85", "W18", "W51", "W94", "W05", "W09", "W01", "W02", "W74", "W41", "WR4", "W06",
                "XRC", "N11", "T18", "N02",
                "UDP", "UR2",
                "Z1-", "ZR0", "Z1TX", "Z1L"]

    found_patterns = []

    # Check for all the patterns in the text and store them
    for pattern in patterns:
        #matches = re.findall(rf"{pattern}[\w-]+", text)
        matches = re.findall(rf"{pattern}[\w-]+(?:\.\d)?", text)
        if matches:
            found_patterns.extend(matches)
    if len(found_patterns) == 0:
        print(text)
    # Handle the case when both LGD occurrences are present
    lgd_matches = [match for match in found_patterns if "LGD" in match]
    found_patterns = [pat.replace(")", "") for pat in found_patterns]
    print(lgd_matches)
    print(found_patterns)
    process_string = lambda input_string: input_string[:len(input_string) // 2] if len(
        input_string) > 20 and not input_string[len(input_string) // 2:].startswith('L') else input_string[
                                                                                              len(input_string) // 2:] if len(
        input_string) > 20 else input_string

    # if len(lgd_matches) > 1:
    #     data["Kama SKU"] = process_string(lgd_matches[0])
    # elif len(found_patterns) > 1:
    #     num = 0
    #     while len(found_patterns[num]) < 10 and num < len(found_patterns):
    #         num = num + 1
    #     data["Kama SKU"] = process_string(found_patterns[num])
    # elif found_patterns:
    #     num1 = 0
    #     while len(found_patterns[num1]) < 10 and num1 < len(found_patterns):
    #         num1 = num1 + 1
    #     data["Kama SKU"] = process_string(found_patterns[num1])
    # else:
    #     data["Kama SKU"] = "Not Found"

    # Extract Metal KT, Metal Color, Size, Req Date, and Qty

    selected_value = None

    if len(lgd_matches) > 1:
        selected_value = lgd_matches[0]

    elif len(found_patterns) > 1:
        for pattern in found_patterns:
            if len(pattern) >= 10:
                selected_value = pattern
                break

    elif found_patterns:
        for pattern in found_patterns:
            if len(pattern) >= 10:
                selected_value = pattern
                break

    # ✅ Final fallback: pick longest from found_patterns
    if not selected_value and found_patterns:
        selected_value = max(found_patterns, key=len)

    # Final assignment
    if selected_value:
        data["Kama SKU"] = process_string(selected_value)
    else:
        data["Kama SKU"] = "Not Found"

    metal_kt_match = re.search(r"\b(\d{2}KT|PLAT|SILV|STER)\b", text)  # Match either Metal KT (e.g., 14KT) or PLAT
    if metal_kt_match:
        try:
            # Start from the position where Metal KT or PLAT ends
            start_pos = metal_kt_match.end()
            # Find the next two space-separated strings for Metal Color and Size
            next_parts = re.findall(r"\S+", text[start_pos:])

            # Assign values with error handling
            data["Metal KT"] = metal_kt_match.group(1) if metal_kt_match else ""
            data["Metal Color"] = next_parts[0] if len(next_parts) > 0 else ""

            # Clean and extract a valid float from the second part (size)
            size_raw = next_parts[1] if len(next_parts) > 1 else ""
            size_raw = correct_double_decimal_number(size_raw)

            # Match a valid number pattern at the beginning (like 55, 55.0, 7.25)
            size_match = re.match(r"^\d+(?:\.\d+)?", size_raw)

            if size_match:
                data["Size"] = f"{float(size_match.group()):.2f}"
            else:
                data["Size"] = ""
            # data["Size"] = f"{float(next_parts[1]):.2f}" if len(next_parts) > 1 and next_parts[1].replace('.', '',
            #                                                                                               1).isdigit() else ""
        except Exception:
            data["Metal KT"] = ""
            data["Metal Color"] = "NA"
            data["Size"] = ""
    else:
        # Handle when no match is found
        data["Metal KT"] = ""
        data["Metal Color"] = "NA"
        data["Size"] = ""

    # Check if the line contains either KT or PLAT
    if "KT" in text or "PLAT" or "SILV" in text:
        # Extract Req Date and Qty (the date is followed by a space and then the quantity)
#*-------------------------------- PW commented / added ------------------------------------------------------------*
        # date_qty_match = re.search(r"(\w{3}/\d{2}/\d{4})\s*(\d+\.?\d*)", text)
        # if date_qty_match:
        #     data["Ship Date"] = datetime.strptime(date_qty_match.group(1), "%b/%d/%Y").strftime("%d-%m-%Y")  # Required Date (e.g., Jan/21/2025)
        #     data["Qty"] = date_qty_match.group(2)  # Quantity (e.g., 1.0)

        date_qty_match = re.search(r"([A-Za-z]{3,9}/\d{2}/\d{4})\s*(\d+\.?\d*)", text)

        if date_qty_match:
            raw_date = date_qty_match.group(1)
            qty = date_qty_match.group(2)

            # Try both short and full month formats
            for fmt in ("%b/%d/%Y", "%B/%d/%Y"):
                try:
                    parsed_date = datetime.strptime(raw_date, fmt)
                    break
                except ValueError:
                    continue
            else:
                raise ValueError(f"Unrecognized date format: {raw_date}")

            data["Ship Date"] = parsed_date.strftime("%d-%m-%Y")
            data["Qty"] = qty

# *--------------------------------------------------------------------------------------------------------------------*

    dia_quality_match = re.search(rf"{re.escape(data['Metal Color'])}\s+(\S+)\s+([A-Z0-9+/-]+(?:\s[A-Z0-9/]+)?)", text)
    #dia_quality_match = re.search(rf"\b{re.escape(data['Metal Color'])}\b\s+(\S+)\s+([A-Z0-9+/-]+(?:\s[A-Z0-9/]+)?)", remaining_text)
    try:

        if dia_quality_match:
            first_part = dia_quality_match.group(1).strip()  # First extracted value after Metal Color
            second_part = dia_quality_match.group(2).strip()  # Next value after first one

            # Try to extract a valid float from the first part (e.g., 55.0.000 → 55.00)
            first_part_match = re.match(r"^\d+(?:\.\d+)?", first_part)

            if first_part_match:
                # It's a number, so second part is likely the diamond quality
                first_part_clean = f"{float(first_part_match.group()):.2f}"
                data["Dia Quality"] = second_part
                wrong_value = second_part
            else:
                # First part is not numeric — treat it as diamond quality
                data["Dia Quality"] = first_part
                wrong_value = first_part

            # Fallback: If extracted Diamond Quality is too long (>10 characters) or too short (<2 characters)
            if len(data["Dia Quality"]) > 10 or len(data["Dia Quality"]) < 3:
                # Find all occurrences of Metal Color in the text
                metal_color_positions = [m.start() for m in re.finditer(re.escape(data["Dia Quality"]), text)]

                if len(metal_color_positions) >= 2:  # If Metal Color appears twice
                    second_metal_color_pos = metal_color_positions[1]  # Get the second occurrence position

                    # Extract the part of the line after the second Metal Color occurrence
                    adjusted_text = text[second_metal_color_pos + len(data["Metal Color"]):]

                    # Look for the correct Diamond Quality format (e.g., "F-VS2", "G+/VS")
                    new_match = re.search(r"([A-Z0-9+/]*[-/][A-Z0-9+/]*)", adjusted_text)

                    if new_match:
                        fallback_value = new_match.group(1).strip()
                        if fallback_value not in data["Dia Quality"]:  # Ensure it's a new value
                            data["Dia Quality"] = fallback_value  # Replace with fallback value
                elif len(metal_color_positions) == 1:
                    first_metal_color_pos = metal_color_positions[0]  # Get the first occurrence position

                    # Extract the part of the line after the first Metal Color occurrence
                    adjusted_text = text[first_metal_color_pos + len(data["Metal Color"]):]

                    size_position = adjusted_text.find(data["Size"]) if data["Size"] in adjusted_text else -1

                    if size_position != -1:
                        # Extract text after Size
                        adjusted_text = adjusted_text[size_position + len(data["Size"]):]

                        # Look for "RD" in the adjusted text
                        rd_position = adjusted_text.find("RD")
                        if rd_position != -1:
                            adjusted_text = adjusted_text[:rd_position]  # Keep only text before "RD"

                        # Look for the correct Diamond Quality format (e.g., "F-VS2", "G+/VS")
                        new_match = re.search(r"([A-Z0-9+/]*[-/][A-Z0-9+/]*)", adjusted_text)

                        if new_match:
                            fallback_value = new_match.group(1).strip()
                            if fallback_value not in data["Dia Quality"]:  # Ensure it's a new value
                                data["Dia Quality"] = fallback_value  # Replace with fallback value
        else:
            data["Dia Quality"] = ""
    except:
        data["Dia Quality"] = ""

    # # Extract ItemType (value after "ItemType")
    # itemtype_match = re.search(r"ItemType\s*(\S+)", text)
    # if itemtype_match:
    #     data['ItemType'] = itemtype_match.group(1)

    # itemtype_match = re.search(
    #     r"(?:ItemType|Item\s+Type)\s+([A-Za-z][\w/&-]*)(?=\s+(Spec|Engraving|Instruction|$))",
    #     text,
    #     re.IGNORECASE
    # )
    itemtype_match = re.search(r"(?:ItemType|Item\s+Type)\s+(\S+)", text, re.IGNORECASE)
    if itemtype_match:
        data["ItemType"] = itemtype_match.group(1).strip()
    else:
        data["ItemType"] = ""

    #engraving_text_match = re.search(r"Engraving Text:\s*(.*?)(?=\s*\*)", text)
    engraving_text_match = re.search(r"Engraving Text:\s*(.*)", text)
    if engraving_text_match:
        data['Engraving'] = engraving_text_match.group(1).strip()
    else:
        data['Engraving'] = ""

    # Extract the text between "Desc." and "GM Min"
    desc_match = re.search(r"Desc\.\s*(.*?)\s*GM Min", text)
    try:
        if desc_match:
            data['Desc'] = desc_match.group(1)

            # Check if "Class" is in the description
            if "Class" in data['Desc']:
                # Remove everything from "Class" onward, and then extract the last word
                cleaned_desc = data['Desc'].split("Class")[0].strip()
            else:
                # If no "Class" is present, just use the whole description
                cleaned_desc = data['Desc']

            # Extract the category (last word in the cleaned description)
            category_match = re.search(r"\b(\w+)$", cleaned_desc)
            if category_match:
                category = category_match.group(1)
                data['Category'] = category
    except:
            data['Category'] = None

    # Extract metal weight range from "GM Min Wt ... Max Wt ..."
    metal_match = re.search(r"GM\s+Min Wt\s+([\d.]+)\s+Max Wt\s+([\d.]+)", text)
    if metal_match:
        metal_wt = f"{metal_match.group(1)}-{metal_match.group(2)}"
    else:
        metal_wt = None

    # Extract diamond weight range from "Diam Min Wt ... Max Wt ..."
    dia_match = re.search(r"Diam\s+Min Wt\s+([\d.]+)\s+Max Wt\s+([\d.]+)", text)
    if dia_match:
        dia_wt = f"{dia_match.group(1)}-{dia_match.group(2)}"
    else:
        dia_wt = None

    # Add the line number to the data
    data["Line No"] = line_no
    data['Metal Tol.'] = metal_wt
    data['Dia Tol.'] = dia_wt

    return data


def extract_data_from_text(text):
    """Extract structured data from text using regex and string manipulation."""
    data = {}

    # Find all matches between 'Purchase From :' and 'Vendor'
    purchase_from_matches = re.search(r"Purchase From\s*:\s*(.*?)\s+Vendor", text)
    if purchase_from_matches:
        data['Vendor'] = purchase_from_matches.group(1)

    # Extract PO Number
    po_match = re.search(r"PO #\s*:\s*(\d+)", text)
    if po_match:
        data["PO Number"] = po_match.group(1)

    print(po_match.group(1))

    # Extract PO Date
    po_date_match = re.search(r"Date\s*:\s*(\w+/\d+/\d+)", text)
    if po_date_match:
        data["PO Date"] = datetime.strptime(po_date_match.group(1), "%b/%d/%Y").strftime("%d-%m-%Y")

    # Extract Vendor PO Number
    # vendor_po_match = re.search(r"Vendor PO #\s*(\w+)", text)
    vendor_po_match = re.search(r"Vendor PO #\s*([^\s]+)", text)
    if vendor_po_match:
        data["Vendor PO"] = vendor_po_match.group(1)

    # Extract Customer Name with prefix "Cust Name"
    cust_name_match = re.search(r"Cust Name\s+(.*?)\s+Gold Price", text)
    if cust_name_match:
        data["Customer Name"] = cust_name_match.group(1).strip()

    # Split the text at "102" if present
    data_list = []
    # Check if 101 and 102 are in the same line

    # # Split the text into lines
    # lines = text.splitlines()
    #
    # # **Always process 101** (as it is always present)
    # if lines[0].startswith("101"):
    #     data_list.append(process_text(lines[0], data.copy(), line_no='101'))
    #     lines.pop(0)  # Remove the processed 101 line
    #
    # # Check if 102 exists at the start of a line, and process dynamically
    # if any(line.startswith("102") for line in lines):
    #     # Find the line where 102 starts
    #     for i, line in enumerate(lines):
    #         if line.startswith("102"):
    #             before_102 = "\n".join(lines[:i])  # All lines before 102
    #             after_102 = "\n".join(lines[i:])  # All lines from 102 onwards
    #             # Process data before 102 and after 102
    #             data_list.append(process_text(before_102, data.copy(), line_no='101'))  # Data before 102
    #             data_list.append(process_text(after_102, data.copy(), line_no='102'))  # Data from 102 onwards
    #             break  # Stop after processing 102
    # else:
    #     # If 102 is not found, treat all remaining lines as part of 101 data
    #     remaining_data = "\n".join(lines)
    #     data_list.append(process_text(remaining_data, data.copy(), line_no='101'))
    #
    # # Check if 103 exists at the start of a line, and process dynamically
    # if any(line.startswith("103") for line in lines):
    #     # Find the line where 103 starts
    #     for i, line in enumerate(lines):
    #         if line.startswith("103"):
    #             before_103 = "\n".join(lines[:i])  # All lines before 103
    #             after_103 = "\n".join(lines[i:])  # All lines from 103 onwards
    #             # Process data before 103 and after 103
    #             #data_list.append(process_text(before_103, data.copy(), line_no='103'))  # Data before 103
    #             data_list.append(process_text(after_103, data.copy(), line_no='103'))  # Data from 103 onwards
    #             break  # Stop after processing 103

    # Loop through all lines and detect any line starting with a number >= 101

    # 🔁 Dynamic Line Processing (for 101, 102, ..., 199+)
    lines = text.splitlines()
    data_list = []
    line_buffer = ""
    current_line_number = None

    # for line in lines:
    #     match = re.match(r"^(\d{3})\s", line)  # Match any 3-digit number
    #     if match:
    #         num = int(match.group(1))
    #         if 101 <= num <= 180:
    #             # Process previous block before starting new one
    #             if line_buffer and current_line_number:
    #                 data_list.append(process_text(line_buffer, data.copy(), line_no=current_line_number))
    #             current_line_number = match.group(1)
    #             line_buffer = line  # Start new block
    #             continue
    #
    #     line_buffer += "\n" + line  # Continue buffering

    def is_block_start(line):
        """Check if line starts with a valid block number (101–180), even if glued"""
        match = re.match(r"^(\d{3})", line)
        if match:
            num = int(match.group(1))
            if 101 <= num <= 400:
                return num
        return None

    for line in lines:
        block_num = is_block_start(line)
        if block_num is not None:
            # Start a new block
            if line_buffer and current_line_number:
                data_list.append(process_text(line_buffer, data.copy(), line_no=current_line_number))
            current_line_number = str(block_num)
            line_buffer = line
            continue

        # Continue previous block
        line_buffer += "\n" + line

    # After the loop, don't forget to process the final buffer
    if line_buffer and current_line_number:
        data_list.append(process_text(line_buffer, data.copy(), line_no=current_line_number))

    return data_list  # Return list of dictionaries for multiple rows


def summary():
    """Main function to process PDFs in a folder and save their structured data."""
    current_dir = os.getcwd()
    mail_download = os.path.join(current_dir, "Mail Download")
    output_folder = os.path.join(current_dir, "Python Output")

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # List PDF files
    pdf_files = list_pdf_files(mail_download)

    final_df = pd.DataFrame()  # Initialize an empty dataframe to store all the data

    # Process each PDF file one at a time
    for pdf_file in pdf_files:
        pdf_path = os.path.join(mail_download, pdf_file)
        output_text_file_name = f"{os.path.splitext(pdf_file)[0]}_text.txt"
        output_text_file_path = os.path.join(output_folder, output_text_file_name)

        # Extract structured data from the current PDF
        extracted_data = extract_data_from_pdf(pdf_path, output_text_file_path)
        # If data is extracted, convert it to a DataFrame and append it to final_df
        if extracted_data:
            # Since the extract_data_from_pdf function returns a list of data (multiple rows),
            # we need to append the data accordingly
            temp_df = pd.DataFrame(extracted_data)  # Convert the extracted data to a DataFrame
            final_df = pd.concat([final_df, temp_df], ignore_index=True)  # Append to the final dataframe
    # After processing all PDFs, save the final dataframe to an Excel file
    current_datetime = datetime.now().strftime("%d.%m.%Y")
    output_excel_file_path = os.path.join(current_dir, "Python Output", f"final_output_data {current_datetime}.xlsx")
    mail_detail_df = pd.read_excel("Mail_Attachments.xlsx", sheet_name='Sheet1')
    print(final_df)
    final_df['PO Number'] = final_df['PO Number'].astype(str)
    mail_detail_df['PO Number'] = mail_detail_df['PO Number'].astype(str)
    final_df1 = pd.merge(final_df, mail_detail_df[['PO Number', 'Req Date']], on='PO Number', how='left')
    try:
        final_df1.to_excel(output_excel_file_path, index=False)
        print(f"Saved all data to {output_excel_file_path}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")

summary()
