"""
Test Integration Script
======================

This script tests the complete PDF to SAP integration workflow.
"""

import os
import logging
from sap_integration import SAPConnection, get_sap_connection_params
from sap_table_operations import SAPTableOperations

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def test_sap_connection():
    """Test SAP connection."""
    logger.info("Testing SAP connection...")
    
    connection_params = get_sap_connection_params()
    sap_conn = SAPConnection(connection_params)
    
    if sap_conn.connect():
        logger.info("✅ SAP connection successful")
        
        if sap_conn.test_connection():
            logger.info("✅ SAP connection test passed")
        else:
            logger.error("❌ SAP connection test failed")
        
        sap_conn.disconnect()
        return True
    else:
        logger.error("❌ SAP connection failed")
        return False


def test_table_operations():
    """Test SAP table operations."""
    logger.info("Testing SAP table operations...")
    
    connection_params = get_sap_connection_params()
    
    try:
        from pyrfc import Connection
        connection = Connection(**connection_params)
        table_ops = SAPTableOperations(connection)
        
        # Test reading a standard table
        test_data = table_ops.read_table('MARA', fields=['MATNR'], max_rows=3)
        if test_data:
            logger.info("✅ Table read operation successful")
        else:
            logger.warning("⚠️ Table read operation returned no data")
        
        # Check if ZKS_T_GB_DEMO exists
        if table_ops.check_table_exists('ZKS_T_GB_DEMO'):
            logger.info("✅ ZKS_T_GB_DEMO table exists")
        else:
            logger.warning("⚠️ ZKS_T_GB_DEMO table does not exist - needs to be created")
        
        connection.close()
        return True
        
    except Exception as e:
        logger.error(f"❌ Table operations test failed: {e}")
        return False


def test_pdf_files():
    """Test if PDF files are available."""
    logger.info("Testing PDF files availability...")
    
    pdf_folder = "Mail Download"
    if os.path.exists(pdf_folder):
        pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
        if pdf_files:
            logger.info(f"✅ Found {len(pdf_files)} PDF files in {pdf_folder}")
            return True
        else:
            logger.warning(f"⚠️ No PDF files found in {pdf_folder}")
            return False
    else:
        logger.error(f"❌ PDF folder {pdf_folder} does not exist")
        return False


def test_python_environment():
    """Test Python environment."""
    logger.info("Testing Python environment...")
    
    import sys
    python_version = sys.version_info
    logger.info(f"Python version: {python_version.major}.{python_version.minor}.{python_version.micro}")
    
    if python_version.major == 3 and python_version.minor == 10:
        logger.info("✅ Python 3.10 environment detected")
    else:
        logger.warning(f"⚠️ Expected Python 3.10, found {python_version.major}.{python_version.minor}")
    
    # Test required packages
    required_packages = ['pyrfc', 'pandas', 'openpyxl', 'pdfplumber']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
            logger.info(f"✅ {package} is available")
        except ImportError:
            missing_packages.append(package)
            logger.error(f"❌ {package} is missing")
    
    if missing_packages:
        logger.error(f"Missing packages: {missing_packages}")
        return False
    else:
        logger.info("✅ All required packages are available")
        return True


def main():
    """Run all tests."""
    logger.info("=" * 60)
    logger.info("PDF to SAP Integration - Test Suite")
    logger.info("=" * 60)
    
    tests = [
        ("Python Environment", test_python_environment),
        ("PDF Files", test_pdf_files),
        ("SAP Connection", test_sap_connection),
        ("Table Operations", test_table_operations)
    ]
    
    results = {}
    
    for test_name, test_func in tests:
        logger.info(f"\n--- {test_name} Test ---")
        try:
            results[test_name] = test_func()
        except Exception as e:
            logger.error(f"Test {test_name} failed with exception: {e}")
            results[test_name] = False
    
    # Summary
    logger.info("\n" + "=" * 60)
    logger.info("TEST SUMMARY")
    logger.info("=" * 60)
    
    passed = 0
    total = len(results)
    
    for test_name, result in results.items():
        status = "✅ PASS" if result else "❌ FAIL"
        logger.info(f"{test_name}: {status}")
        if result:
            passed += 1
    
    logger.info(f"\nOverall: {passed}/{total} tests passed")
    
    if passed == total:
        logger.info("🎉 All tests passed! Integration is ready to use.")
    else:
        logger.info("⚠️ Some tests failed. Please check the issues above.")
    
    # Next steps
    logger.info("\n" + "=" * 60)
    logger.info("NEXT STEPS")
    logger.info("=" * 60)
    
    if not results.get("Table Operations", False):
        logger.info("1. Create ZKS_T_GB_DEMO table in SAP")
        logger.info("2. Create ZKS_FM_GB_DEMO function module in SAP")
    
    if results.get("SAP Connection", False) and results.get("PDF Files", False):
        logger.info("3. Run: python pdf_to_sap_integration.py")
    else:
        logger.info("3. Fix the issues above before running the integration")


if __name__ == "__main__":
    main()
