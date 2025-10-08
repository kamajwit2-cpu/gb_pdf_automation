"""
SAP Integration Module for GB PDF Automation
============================================

This module handles SAP connectivity and data transfer for the GB PDF automation system.
It connects to SAP using PyRFC and pushes Excel data to the ZKS_T_GB_DEMO table.
"""

import pandas as pd
from pyrfc import Connection, RFCError
from typing import Dict, List, Any, Optional
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class SAPConnection:
    """SAP connection handler using PyRFC."""
    
    def __init__(self, connection_params: Dict[str, str]):
        """
        Initialize SAP connection with provided parameters.
        
        Args:
            connection_params: Dictionary containing SAP connection parameters
        """
        self.connection_params = connection_params
        self.connection = None
        
    def connect(self) -> bool:
        """
        Establish connection to SAP system.
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        try:
            self.connection = Connection(**self.connection_params)
            logger.info("Successfully connected to SAP system")
            return True
        except RFCError as e:
            logger.error(f"SAP connection failed: {e}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error during SAP connection: {e}")
            return False
    
    def disconnect(self):
        """Close SAP connection."""
        if self.connection:
            self.connection.close()
            logger.info("SAP connection closed")
    
    def test_connection(self) -> bool:
        """
        Test the SAP connection by calling a simple RFC.
        
        Returns:
            bool: True if test successful, False otherwise
        """
        try:
            if not self.connection:
                logger.error("No active SAP connection")
                return False
                
            # Test with RFC_SYSTEM_INFO
            result = self.connection.call('RFC_SYSTEM_INFO')
            logger.info(f"SAP system info: {result.get('RFCSI_EXPORT', {}).get('RFCSI_EXPORT', 'Unknown')}")
            return True
        except RFCError as e:
            logger.error(f"SAP connection test failed: {e}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error during SAP connection test: {e}")
            return False


class SAPDataPusher:
    """Handles pushing Excel data to SAP tables."""
    
    def __init__(self, sap_connection: SAPConnection):
        """
        Initialize SAP data pusher.
        
        Args:
            sap_connection: Active SAP connection instance
        """
        self.sap_connection = sap_connection
        
    def push_to_zks_table(self, data: List[Dict[str, Any]], table_name: str = "ZKS_T_GB_DEMO") -> bool:
        """
        Push data to ZKS_T_GB_DEMO table using the custom function module.
        
        Args:
            data: List of dictionaries containing data to push
            table_name: Name of the SAP table (default: ZKS_T_GB_DEMO)
            
        Returns:
            bool: True if push successful, False otherwise
        """
        try:
            if not self.sap_connection.connection:
                logger.error("No active SAP connection")
                return False
            
            # Prepare data for SAP function module
            sap_data = self._prepare_sap_data(data)
            
            # Call the custom function module ZKS_FM_GB_DEMO
            result = self.sap_connection.connection.call(
                'ZKS_FM_GB_DEMO',
                IT_DATA=sap_data,
                IV_TABLE_NAME=table_name
            )
            
            logger.info(f"Successfully pushed {len(data)} records to SAP table {table_name}")
            logger.info(f"SAP response: {result}")
            return True
            
        except RFCError as e:
            logger.error(f"SAP RFC call failed: {e}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error during SAP data push: {e}")
            return False
    
    def _prepare_sap_data(self, data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Prepare data for SAP function module call.
        
        Args:
            data: List of dictionaries containing Excel data
            
        Returns:
            List of dictionaries formatted for SAP
        """
        sap_data = []
        
        for record in data:
            # Map Excel columns to SAP table fields
            sap_record = {
                'VENDOR': record.get('Vendor', ''),
                'PO_NUMBER': record.get('PO Number', ''),
                'VPO': record.get('VPO', ''),
                'CUSTOMER_NAME': record.get('Customer Name', ''),
                'KAMA_SKU': record.get('Kama SKU', ''),
                'METAL_KT': record.get('Metal KT', ''),
                'METAL_COLOR': record.get('Metal Color', ''),
                'SIZE': record.get('Size', ''),
                'SHIP_DATE': record.get('Ship Date', ''),
                'QTY': record.get('Qty', ''),
                'DIA_QUALITY': record.get('Dia Quality', ''),
                'ITEM_TYPE': record.get('ItemType', ''),
                'ENGRAVING': record.get('Engraving', ''),
                'DESC': record.get('Desc', ''),
                'CATEGORY': record.get('Category', ''),
                'LINE_NO': record.get('Line No', ''),
                'METAL_TOL': record.get('Metal Tol.', ''),
                'DIA_TOL': record.get('Dia Tol.', ''),
                'REQ_DATE': record.get('Req Date', ''),
                'CREATED_DATE': datetime.now().strftime('%Y%m%d'),
                'CREATED_TIME': datetime.now().strftime('%H%M%S')
            }
            sap_data.append(sap_record)
        
        return sap_data
    
    def push_excel_to_sap(self, excel_file_path: str, sheet_name: str = None) -> bool:
        """
        Read Excel file and push data to SAP.
        
        Args:
            excel_file_path: Path to the Excel file
            sheet_name: Name of the sheet to read (if None, reads first sheet)
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Read Excel file
            if sheet_name:
                df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(excel_file_path)
            
            # Convert DataFrame to list of dictionaries
            data = df.to_dict('records')
            
            logger.info(f"Read {len(data)} records from Excel file: {excel_file_path}")
            
            # Push to SAP
            return self.push_to_zks_table(data)
            
        except Exception as e:
            logger.error(f"Error reading Excel file or pushing to SAP: {e}")
            return False


def get_sap_connection_params() -> Dict[str, str]:
    """
    Get SAP connection parameters from the PyRFC.py file.
    
    Returns:
        Dictionary containing SAP connection parameters
    """
    return {
        'ashost': '192.168.150.72',
        'sysnr': '00',
        'client': '333',
        'user': 'ksbot',
        'passwd': 'ksbot@123',
        'lang': 'EN'
    }


def main():
    """Main function to test SAP integration."""
    logger.info("Starting SAP integration test...")
    
    # Get connection parameters
    connection_params = get_sap_connection_params()
    
    # Create and test SAP connection
    sap_conn = SAPConnection(connection_params)
    
    if sap_conn.connect():
        logger.info("SAP connection established successfully")
        
        # Test connection
        if sap_conn.test_connection():
            logger.info("SAP connection test passed")
            
            # Create data pusher
            data_pusher = SAPDataPusher(sap_conn)
            
            # Test with sample data
            sample_data = [
                {
                    'Vendor': 'TEST_VENDOR',
                    'PO Number': '12345',
                    'VPO': 'TEST_VPO',
                    'Customer Name': 'TEST_CUSTOMER',
                    'Kama SKU': 'TEST_SKU',
                    'Metal KT': '14KT',
                    'Metal Color': 'W',
                    'Size': '7.00',
                    'Ship Date': '01-01-2025',
                    'Qty': '1',
                    'Dia Quality': 'F-VS2',
                    'ItemType': 'TEST',
                    'Engraving': '',
                    'Desc': 'Test Description',
                    'Category': 'RING',
                    'Line No': '101',
                    'Metal Tol.': '1.0-2.0',
                    'Dia Tol.': '0.5-1.0',
                    'Req Date': '01-01-2025'
                }
            ]
            
            # Push sample data
            if data_pusher.push_to_zks_table(sample_data):
                logger.info("Sample data pushed successfully")
            else:
                logger.error("Failed to push sample data")
        else:
            logger.error("SAP connection test failed")
    else:
        logger.error("Failed to establish SAP connection")
    
    # Disconnect
    sap_conn.disconnect()
    logger.info("SAP integration test completed")


if __name__ == "__main__":
    main()
