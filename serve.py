"""
Production Server for GB PDF Automation
======================================

This module serves the Flask application using Waitress WSGI server
for production deployment with multiple concurrent users.
"""

import os
import sys
from waitress import serve
from app import app

def main():
    """Start the production server."""
    
    # Set production environment
    os.environ['ENV'] = 'production'
    
    # Server configuration
    host = '0.0.0.0'  # Bind to all interfaces for local network access
    port = 5002
    threads = 4  # Number of threads for concurrent users
    
    print("=" * 60)
    print("GB PDF Automation - Production Server")
    print("=" * 60)
    print(f"Starting server on http://{host}:{port}")
    print(f"Environment: {os.environ.get('ENV', 'development')}")
    print(f"Threads: {threads}")
    print("=" * 60)
    print("Server is running... Press Ctrl+C to stop")
    print("=" * 60)
    
    try:
        # Start Waitress server
        serve(
            app,
            host=host,
            port=port,
            threads=threads,
            url_scheme='http',
            ident='GB-PDF-Automation/1.0'
        )
    except KeyboardInterrupt:
        print("\n" + "=" * 60)
        print("Server stopped by user")
        print("=" * 60)
    except Exception as e:
        print(f"\nError starting server: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
