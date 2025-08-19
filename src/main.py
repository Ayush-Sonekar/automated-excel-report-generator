#!/usr/bin/env python3
"""
Automated Excel Report Generator
Main entry point for the application that processes business data and generates professional Excel reports.
"""

import sys
import os
import argparse
from datetime import datetime
from data_processor import DataProcessor
from report_generator import ReportGenerator
from data_intelligence import DataIntelligence
from config import Config

def main():
    """Main function to orchestrate the report generation process."""
    parser = argparse.ArgumentParser(description='Generate professional Excel reports from CSV data')
    parser.add_argument('input_file', nargs='?', default='nyc-rolling-sales.csv', 
                       help='Path to input CSV file (default: sample_data.csv)')
    parser.add_argument('--output', '-o', default=None, 
                       help='Output Excel file path (default: auto-generated with timestamp)')
    parser.add_argument('--verbose', '-v', action='store_true', 
                       help='Enable verbose output')
    
    args = parser.parse_args()
    
    # Validate input file exists
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' not found.")
        print("Please provide a valid CSV file path or use the default 'sample_data.csv'")
        return 1
    
    # Generate output filename with timestamp if not provided
    if args.output is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        args.output = f"Monthly_Sales_Report_{timestamp}.xlsx"
    
    try:
        if args.verbose:
            print(f"Processing data from: {args.input_file}")
            print(f"Output will be saved to: {args.output}")
        
        # Initialize components
        processor = DataProcessor(verbose=args.verbose)
        intelligence = DataIntelligence(verbose=args.verbose)
        
        # Load and process data
        if args.verbose:
            print("Loading and cleaning data...")
        raw_data = processor.load_data(args.input_file)
        
        if raw_data.empty:
            print("Error: No data found in the input file or all data was filtered out.")
            return 1
        
        # Smart data analysis
        if args.verbose:
            print("Analyzing data type and generating insights...")
        domain_info = intelligence.detect_data_type(raw_data)
        if args.verbose:
            print(f"Detected data type: {domain_info['domain']} (confidence: {domain_info['confidence']:.1f}%)")
        
        # Process and aggregate data
        if args.verbose:
            print("Processing and aggregating data...")
        processed_data = processor.process_data(raw_data)
        
        # Generate intelligent insights
        insights = intelligence.generate_insights(processed_data['raw_data'], processed_data['aggregations'], domain_info)
        processed_data['domain_info'] = domain_info
        processed_data['insights'] = insights
        
        # Generate report
        if args.verbose:
            print("Generating Excel report...")
        report_gen = ReportGenerator(verbose=args.verbose)
        success = report_gen.generate_report(processed_data, args.output)
        
        if success:
            print(f"Report successfully generated: {args.output}")
            print(f"Report contains {len(processed_data['raw_data'])} data records")
            return 0
        else:
            print("Error: Failed to generate report")
            return 1
            
    except Exception as e:
        print(f"Error: {str(e)}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
