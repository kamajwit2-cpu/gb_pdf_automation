# PDF to SAP Integration - Complete Setup Guide

## Overview
This project integrates PDF data extraction with SAP data push functionality. It processes PDF files containing purchase order data and pushes the extracted information to a custom SAP table.

## Current Status
✅ **SAP Connection**: Successfully established  
✅ **PDF Extraction**: Working with existing system  
✅ **Python 3.10 Environment**: Created and configured  
❌ **SAP Table**: `ZKS_T_GB_DEMO` - Not found (needs to be created)  
❌ **Function Module**: `ZKS_FM_GB_DEMO` - Not found (needs to be created)  

## SAP Connection Details
- **Host**: 192.168.150.72
- **System Number**: 00
- **Client**: 333
- **User**: ksbot
- **Password**: ksbot@123
- **Language**: EN

## What You Need to Do in SAP

### 1. Create the Table ZKS_T_GB_DEMO
You need to create a custom table in SAP with the following structure:

```abap
Table: ZKS_T_GB_DEMO
Fields:
- MANDT (Client) - CHAR 3
- VENDOR - CHAR 10
- PO_NUMBER - CHAR 10
- VPO - CHAR 50
- CUSTOMER_NAME - CHAR 50
- KAMA_SKU - CHAR 30
- METAL_KT - CHAR 10
- METAL_COLOR - CHAR 5
- SIZE - CHAR 10
- SHIP_DATE - CHAR 10
- QTY - CHAR 10
- DIA_QUALITY - CHAR 20
- ITEM_TYPE - CHAR 20
- ENGRAVING - CHAR 100
- DESCRIPTION - CHAR 200
- CATEGORY - CHAR 20
- LINE_NO - CHAR 10
- METAL_TOL - CHAR 20
- DIA_TOL - CHAR 20
- REQ_DATE - CHAR 10
- CREATED_DATE - CHAR 8
- CREATED_TIME - CHAR 6
- CREATED_BY - CHAR 12
```

### 2. Create Function Module ZKS_FM_GB_DEMO
Create a function module with the following interface:

```abap
Function Module: ZKS_FM_GB_DEMO

Import Parameters:
- IT_DATA (Type: Table of ZKS_T_GB_DEMO)
- IV_TABLE_NAME (Type: CHAR 30)

Export Parameters:
- EV_SUCCESS (Type: CHAR 1)
- EV_MESSAGE (Type: CHAR 255)

Tables:
- IT_DATA (Type: ZKS_T_GB_DEMO)

Function Module Code:
FUNCTION ZKS_FM_GB_DEMO.
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  IMPORTING
*"     VALUE(IV_TABLE_NAME) TYPE CHAR30
*"  TABLES
*"     IT_DATA TYPE ZKS_T_GB_DEMO
*"  EXPORTING
*"     VALUE(EV_SUCCESS) TYPE CHAR1
*"     VALUE(EV_MESSAGE) TYPE CHAR255
*"----------------------------------------------------------------------

  DATA: lv_count TYPE i.
  
  CLEAR: ev_success, ev_message.
  
  LOOP AT it_data.
    INSERT zks_t_gb_demo FROM it_data.
    IF sy-subrc = 0.
      lv_count = lv_count + 1.
    ENDIF.
  ENDLOOP.
  
  IF lv_count > 0.
    ev_success = 'X'.
    ev_message = |Successfully inserted { lv_count } records|.
  ELSE.
    ev_success = ''.
    ev_message = 'No records were inserted'.
  ENDIF.

ENDFUNCTION.
```

## Files Created

### 1. `sap_integration.py`
- Core SAP connectivity module
- Handles connection management
- Provides data push functionality

### 2. `pdf_to_sap_integration.py`
- Complete integration workflow
- Combines PDF extraction with SAP push
- Main script to run the entire process

### 3. `sap_table_operations.py`
- Alternative SAP table operations
- Table structure inspection
- Standard RFC operations

### 4. `requirements_sap.txt`
- Python dependencies for SAP integration
- PyRFC and related packages

## How to Use

### 1. Activate the Python 3.10 Environment
```bash
.\sap_integration_env\Scripts\Activate.ps1
```

### 2. Run the Complete Integration
```bash
python pdf_to_sap_integration.py
```

### 3. Test SAP Connection Only
```bash
python sap_integration.py
```

### 4. Test Table Operations
```bash
python sap_table_operations.py
```

## Expected Workflow

1. **PDF Processing**: Extracts data from PDF files in "Mail Download" folder
2. **Excel Generation**: Creates combined Excel file with all extracted data
3. **SAP Push**: Pushes data to ZKS_T_GB_DEMO table using ZKS_FM_GB_DEMO function module
4. **Logging**: Provides detailed logs of the entire process

## Troubleshooting

### SAP Connection Issues
- Verify network connectivity to 10.81.2.236
- Check SAP user credentials
- Ensure SAP RFC service is running

### Table/Function Module Issues
- Create the table ZKS_T_GB_DEMO in SAP
- Create the function module ZKS_FM_GB_DEMO
- Ensure proper authorization for the ksbot user

### Python Environment Issues
- Use Python 3.10 (required for PyRFC compatibility)
- Install all dependencies from requirements_sap.txt
- Ensure virtual environment is activated

## Next Steps

1. **Create SAP Table**: Set up ZKS_T_GB_DEMO table in SAP
2. **Create Function Module**: Implement ZKS_FM_GB_DEMO function module
3. **Test Integration**: Run the complete workflow
4. **Production Deployment**: Deploy to production environment

## Support

If you encounter any issues:
1. Check the logs for detailed error messages
2. Verify SAP connectivity and permissions
3. Ensure all required SAP objects are created
4. Test individual components before running the full integration
