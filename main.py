# Project 4: Automated Monthly Reporting System
# Author: Nabila Salvaningtyas
# Description: Orchestrates the data generation and report creation process.

import os
import sys

# Add 'src' to the python path so we can import modules from it
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from data_generator import generate_data
from report_generator import generate_report

# Define Constants
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, 'data', 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'data', 'output')
INPUT_FILE = os.path.join(INPUT_DIR, 'raw_sales_data.csv')
OUTPUT_FILE = os.path.join(OUTPUT_DIR, 'Monthly_Report.xlsx')

def main():
    print("=== Automated Monthly Reporting System (Enhanced) ===")
    
    # Ensure directories exist
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Check if input data exists
    if not os.path.exists(INPUT_FILE):
        print(f"Input data not found at {INPUT_FILE}. Generating dummy data...")
        generate_data(INPUT_FILE)
    else:
        print(f"Input data found at {INPUT_FILE}.")
        
    # Generate Report
    print("Starting report generation...")
    generate_report(INPUT_FILE, OUTPUT_FILE)
    
    print("\n=== Process Complete ===")
    print(f"Report generated at: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
