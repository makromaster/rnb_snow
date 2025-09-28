#!/usr/bin/env python3
"""
Installation script for RnB Snow Ticket Matching System
Handles compatibility issues with pandas/numpy versions
"""

import subprocess
import sys

def install_packages():
    """Install required packages with compatible versions"""

    packages = [
        "selenium>=4.0.0"
    ]

    print("Installing RnB Snow Ticket Matching System dependencies...")
    print("Only Selenium is required - other dependencies are built into Python.\n")

    for package in packages:
        print(f"Installing {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"✓ {package} installed successfully")
        except subprocess.CalledProcessError as e:
            print(f"✗ Failed to install {package}: {e}")
            return False
        print()

    print("Installation complete!")
    print("\nTo run the application:")
    print("python gui_main.py")

    return True

def check_installation():
    """Check if packages are properly installed"""
    print("\nVerifying installation...")

    try:
        import selenium
        print(f"✓ selenium {selenium.__version__}")

        import sqlite3
        print("✓ sqlite3 (built-in)")

        import csv
        print("✓ csv (built-in)")

        import tkinter
        print("✓ tkinter (built-in)")

        # Test CSV functionality
        import io
        buffer = io.StringIO()
        writer = csv.writer(buffer)
        writer.writerow(['test', 'data'])
        print("✓ CSV export functionality working")

        print("\n✓ All dependencies verified successfully!")
        return True

    except Exception as e:
        print(f"✗ Verification failed: {e}")
        return False

if __name__ == "__main__":
    print("RnB Snow Ticket Matching System - Dependency Installer")
    print("=" * 55)

    if install_packages():
        check_installation()
    else:
        print("Installation failed. Please check error messages above.")
        sys.exit(1)