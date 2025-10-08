"""
Test 4 Columns Insert
====================

This script tests inserting data into the 4 existing columns of ZKS_T_GB_DEMO table:
- MANDT (Client)
- USER_NAME (User Name) 
- USER_DATE (Date)
- USER_TIME (Time)
"""

import logging
from datetime import datetime
from sap_integration import get_sap_connection_params
from pyrfc import Connection, RFCError

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class Test4ColumnsInsert:
    """Test inserting data into 4 existing columns."""
    
    def __init__(self):
        self.sap_connection = None
        
    def initialize_sap_connection(self) -> bool:
        """Initialize SAP connection."""
        try:
            connection_params = get_sap_connection_params()
            self.sap_connection = Connection(**connection_params)
            logger.info("SAP connection initialized successfully")
            return True
        except Exception as e:
            logger.error(f"Error initializing SAP connection: {e}")
            return False
    
    def create_test_data(self, count: int = 5) -> list:
        """
        Create test data for the 4 existing columns.
        
        Args:
            count: Number of test records to create
            
        Returns:
            List of test records
        """
        test_data = []
        current_time = datetime.now()
        
        for i in range(count):
            # Create time string in HHMMSS format
            time_str = f"{current_time.strftime('%H%M%S')}{i:02d}"
            # Ensure it's exactly 6 characters
            time_str = time_str[:6]
            
            record = {
                'MANDT': '333',  # Client
                'USER_NAME': f'KSBOT_{i+1:03d}',  # User name with sequence
                'USER_DATE': current_time.strftime('%Y%m%d'),  # Current date
                'USER_TIME': time_str  # Time string in HHMMSS format
            }
            test_data.append(record)
        
        logger.info(f"Created {len(test_data)} test records for 4 columns")
        for i, record in enumerate(test_data):
            logger.info(f"  Record {i+1}: {record}")
        
        return test_data
    
    def test_function_module(self, test_data: list) -> bool:
        """
        Test the ZKS_FM_GB_DEMO function module.
        
        Args:
            test_data: List of test records
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            logger.info("Testing ZKS_FM_GB_DEMO function module...")
            logger.info("Using interface: IT_INPUT_ (table parameter)")
            
            result = self.sap_connection.call(
                'ZKS_FM_GB_DEMO',
                IT_INPUT_=test_data
            )
            
            logger.info(f"Function module result: {result}")
            logger.info("✅ SUCCESS: Function module executed without errors")
            return True
                
        except RFCError as e:
            if "FU_NOT_FOUND" in str(e):
                logger.error("❌ Function module ZKS_FM_GB_DEMO not found in SAP")
                logger.info("Please create the function module in SAP first")
            else:
                logger.error(f"❌ Function module error: {e}")
            return False
        except Exception as e:
            logger.error(f"❌ Unexpected error: {e}")
            return False
    
    def verify_insertion(self) -> bool:
        """Verify that data was actually inserted into the table."""
        try:
            logger.info("Verifying data insertion...")
            
            result = self.sap_connection.call(
                'RFC_READ_TABLE',
                QUERY_TABLE='ZKS_T_GB_DEMO',
                ROWCOUNT=10
            )
            
            if 'DATA' in result:
                record_count = len(result['DATA'])
                logger.info(f"Table now has {record_count} records")
                
                if record_count > 0:
                    logger.info("✅ SUCCESS: Data was inserted into the table!")
                    for i, row in enumerate(result['DATA']):
                        logger.info(f"  Record {i+1}: {row['WA']}")
                    return True
                else:
                    logger.error("❌ FAILED: No data found in table")
                    return False
            else:
                logger.error("❌ FAILED: Could not read table data")
                return False
                
        except Exception as e:
            logger.error(f"❌ Error verifying insertion: {e}")
            return False
    
    def run_test(self) -> bool:
        """Run the complete test."""
        logger.info("Starting 4 Columns Insert Test")
        logger.info("=" * 50)
        
        # Initialize connection
        if not self.initialize_sap_connection():
            return False
        
        try:
            # Create test data
            test_data = self.create_test_data(5)
            
            # Test function module
            if self.test_function_module(test_data):
                # Verify insertion
                if self.verify_insertion():
                    logger.info("🎉 COMPLETE SUCCESS: 4 columns test passed!")
                    return True
                else:
                    logger.error("❌ Insertion verification failed")
                    return False
            else:
                logger.error("❌ Function module test failed")
                return False
                
        finally:
            if self.sap_connection:
                self.sap_connection.close()
                logger.info("SAP connection closed")
    
    def close_connection(self):
        """Close SAP connection."""
        if self.sap_connection:
            self.sap_connection.close()


def main():
    """Main function to run the test."""
    logger.info("4 Columns Insert Test")
    logger.info("=" * 30)
    
    tester = Test4ColumnsInsert()
    
    try:
        success = tester.run_test()
        
        if success:
            logger.info("\n🎉 TEST PASSED: 4 columns insert working!")
            logger.info("Next step: Add more columns to table or function module")
        else:
            logger.info("\n⚠️ TEST FAILED: Need to create function module in SAP")
            logger.info("Next step: Create ZKS_FM_GB_DEMO function module")
            
    finally:
        tester.close_connection()
    
    logger.info("4 columns test completed")


if __name__ == "__main__":
    main()
