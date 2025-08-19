"""
Data Processing Module
Handles data loading, cleaning, validation, and aggregation for the report generator.
"""

import pandas as pd
import numpy as np
from datetime import datetime
from config import Config

class DataProcessor:
    """Handles all data processing operations including cleaning and aggregation."""
    
    def __init__(self, verbose=False):
        self.verbose = verbose
        self.config = Config()
    
    def load_data(self, file_path):
        """
        Load data from CSV file with error handling and validation.
        
        Args:
            file_path (str): Path to the CSV file
            
        Returns:
            pd.DataFrame: Loaded data
        """
        try:
            # Try different encodings if default fails
            encodings = ['utf-8', 'latin-1', 'cp1252']
            df = None
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    if self.verbose:
                        print(f"Successfully loaded data with {encoding} encoding")
                    break
                except UnicodeDecodeError:
                    continue
            
            if df is None:
                raise ValueError("Unable to read CSV file with any supported encoding")
            
            if self.verbose:
                print(f"Loaded {len(df)} rows and {len(df.columns)} columns")
                print(f"Columns: {list(df.columns)}")
            
            return df
            
        except Exception as e:
            raise Exception(f"Error loading data from {file_path}: {str(e)}")
    
    def clean_data(self, df):
        """
        Clean the dataset by handling missing values, duplicates, and data types.
        
        Args:
            df (pd.DataFrame): Raw dataframe
            
        Returns:
            pd.DataFrame: Cleaned dataframe
        """
        if self.verbose:
            print(f"Starting data cleaning. Initial shape: {df.shape}")
        
        # Create a copy to avoid modifying original data
        cleaned_df = df.copy()
        
        # Process full dataset - no sampling limitations
        if self.verbose:
            print(f"Processing full dataset: {len(cleaned_df):,} rows")
        
        # Enhanced cleaning for real estate data with " - " values
        for col in cleaned_df.columns:
            if cleaned_df[col].dtype == 'object':
                # Replace " - " and similar null indicators with proper NaN
                cleaned_df[col] = cleaned_df[col].replace([' - ', ' -  ', '-', '  ', '', 'N/A'], pd.NA)
        
        # Handle missing values
        initial_rows = len(cleaned_df)
        
        # Special handling for price columns in real estate data
        price_columns = [col for col in cleaned_df.columns if any(keyword in col.lower() for keyword in ['price', 'sale'])]
        for col in price_columns:
            if col in cleaned_df.columns and cleaned_df[col].dtype == 'object':
                # Convert price strings to numeric, handling commas and invalid values
                cleaned_df[col] = pd.to_numeric(
                    cleaned_df[col].astype(str).str.replace(',', '').str.replace(' ', '').str.replace('-', '').replace('', '0'),
                    errors='coerce'
                )
                if self.verbose:
                    print(f"Converted {col} to numeric (price column)")
        
        # For numeric columns, handle missing values intelligently
        numeric_columns = cleaned_df.select_dtypes(include=[np.number]).columns
        for col in numeric_columns:
            if cleaned_df[col].isnull().sum() > 0:
                # For price columns, don't fill with median - keep as NaN for proper exclusion
                if any(keyword in col.lower() for keyword in ['price', 'sale', 'value', 'amount']):
                    if self.verbose:
                        print(f"Keeping {col} missing values as NaN for accurate calculations (preserves data integrity)")
                else:
                    # For other numeric columns, fill with median
                    median_val = cleaned_df[col].median()
                    if pd.isna(median_val) or median_val == 0:
                        median_val = cleaned_df[col].mean()
                        if pd.isna(median_val):
                            median_val = 0
                    cleaned_df = cleaned_df.copy()
                    cleaned_df[col] = cleaned_df[col].fillna(median_val)
                    if self.verbose:
                        print(f"Filled {col} missing values with median: {median_val:.2f}")
        
        # For categorical columns, fill with mode or 'Unknown'
        categorical_columns = cleaned_df.select_dtypes(include=['object']).columns
        for col in categorical_columns:
            if cleaned_df[col].isnull().sum() > 0:
                if not cleaned_df[col].mode().empty:
                    mode_val = cleaned_df[col].mode()[0]
                    cleaned_df[col] = cleaned_df[col].fillna(mode_val)
                else:
                    cleaned_df[col].fillna('Unknown', inplace=True)
                if self.verbose:
                    print(f"Filled {col} missing values")
        
        # Remove duplicates
        duplicates_before = len(cleaned_df)
        cleaned_df.drop_duplicates(inplace=True)
        duplicates_removed = duplicates_before - len(cleaned_df)
        
        if self.verbose:
            print(f"Removed {duplicates_removed} duplicate rows")
            print(f"Final cleaned shape: {cleaned_df.shape}")
        
        return cleaned_df
    
    def standardize_columns(self, df):
        """
        Standardize column names and detect common business data patterns.
        
        Args:
            df (pd.DataFrame): Input dataframe
            
        Returns:
            pd.DataFrame: Dataframe with standardized columns
        """
        standardized_df = df.copy()
        
        # Common column name mappings
        column_mappings = {
            'date': ['date', 'order_date', 'transaction_date', 'sale_date'],
            'sales': ['sales', 'amount', 'revenue', 'total', 'value','sale_price'],
            'product': ['product', 'item', 'product_name', 'item_name','building_class_at_present'],
            'region': ['region', 'location', 'area', 'territory','zip_code'],
            'month': ['month', 'month_name','sale_date'],
            'quantity': ['quantity', 'qty', 'units', 'count','total_units']
        }
        
        # Normalize column names (lowercase, remove spaces)
        standardized_df.columns = [col.lower().strip().replace(' ', '_') for col in standardized_df.columns]
        
        # Try to identify and standardize key columns
        for standard_name, possible_names in column_mappings.items():
            for col in standardized_df.columns:
                if any(possible in col.lower() for possible in possible_names):
                    if standard_name not in standardized_df.columns:
                        standardized_df = standardized_df.rename(columns={col: standard_name})
                        if self.verbose:
                            print(f"Standardized column '{col}' to '{standard_name}'")
                        break
        
        return standardized_df
    
    def parse_dates(self, df):
        """
        Parse and standardize date columns.
        
        Args:
            df (pd.DataFrame): Input dataframe
            
        Returns:
            pd.DataFrame: Dataframe with parsed dates
        """
        date_df = df.copy()
        
        # Common date column names
        date_columns = [col for col in date_df.columns if 'date' in col.lower()]
        
        for col in date_columns:
            try:
                date_df[col] = pd.to_datetime(date_df[col], errors='coerce')
                if self.verbose:
                    print(f"Parsed {col} as datetime")
                
                # Create additional time-based columns for analysis
                if col == 'date' or 'date' in col.lower():
                    date_df['year'] = date_df[col].dt.year
                    date_df['month'] = date_df[col].dt.month
                    date_df['month_name'] = date_df[col].dt.month_name()
                    date_df['quarter'] = date_df[col].dt.quarter
                    
            except Exception as e:
                if self.verbose:
                    print(f"Could not parse {col} as date: {e}")
        
        return date_df
    
    def aggregate_data(self, df):
        """
        Create various data aggregations for reporting.
        
        Args:
            df (pd.DataFrame): Cleaned dataframe
            
        Returns:
            dict: Dictionary containing various aggregated datasets
        """
        aggregations = {}
        
        # Smart sales column detection - prioritize price/value columns over quantity
        sales_col = None
        
        # Priority 1: Look for price/value columns first (better for real estate)
        price_keywords = ['price', 'value', 'amount', 'cost']
        for col in df.columns:
            if any(keyword in col.lower() for keyword in price_keywords):
                if pd.api.types.is_numeric_dtype(df[col]):
                    sales_col = col
                    if self.verbose:
                        print(f"Using {sales_col} as primary value column (price-based)")
                    break
        
        # Priority 2: Look for revenue/sales columns
        if sales_col is None:
            revenue_keywords = ['sales', 'revenue', 'income']
            for col in df.columns:
                if any(keyword in col.lower() for keyword in revenue_keywords):
                    if pd.api.types.is_numeric_dtype(df[col]):
                        sales_col = col
                        if self.verbose:
                            print(f"Using {sales_col} as sales/revenue column")
                        break
        
        # Priority 3: Avoid unit/quantity columns, use meaningful numeric columns
        if sales_col is None:
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            # Filter out unit/quantity columns
            excluded_keywords = ['units', 'quantity', 'count', 'id', 'year', 'month', 'block', 'lot']
            meaningful_cols = [col for col in numeric_cols if not any(keyword in col.lower() for keyword in excluded_keywords)]
            
            if len(meaningful_cols) > 0:
                sales_col = meaningful_cols[0]
                if self.verbose:
                    print(f"Using {sales_col} as primary numeric column")
            elif len(numeric_cols) > 0:
                sales_col = numeric_cols[0]
                if self.verbose:
                    print(f"Fallback: Using {sales_col} as sales column")
        
        if sales_col:
            # Create a filtered dataset with only valid sales for aggregations
            valid_sales_df = df[df[sales_col].notna() & (df[sales_col] > 0)].copy()
            
            if self.verbose:
                print(f"Using {len(valid_sales_df):,} valid sales records (out of {len(df):,} total) for accurate aggregations")
            
            # Monthly aggregation with valid sales only
            if 'month_name' in valid_sales_df.columns and len(valid_sales_df) > 0:
                monthly_agg = valid_sales_df.groupby('month_name')[sales_col].agg(['sum', 'count', 'mean']).reset_index()
                monthly_agg.columns = ['Month', 'Total_Sales', 'Transaction_Count', 'Average_Sale']
                monthly_agg = monthly_agg.round(2)
                aggregations['monthly'] = monthly_agg
            
            # Product aggregation
            product_col = None
            for col in df.columns:
                if any(keyword in col.lower() for keyword in ['product', 'item']):
                    product_col = col
                    break
            
            if product_col and len(valid_sales_df) > 0:
                product_agg = valid_sales_df.groupby(product_col)[sales_col].agg(['sum', 'count', 'mean']).reset_index()
                product_agg.columns = ['Product', 'Total_Sales', 'Units_Sold', 'Average_Price']
                product_agg = product_agg.sort_values('Total_Sales', ascending=False).head(20)
                product_agg = product_agg.round(2)
                aggregations['product'] = product_agg
            
            # Regional aggregation with valid sales only
            region_col = None
            for col in valid_sales_df.columns:
                if any(keyword in col.lower() for keyword in ['region', 'location', 'area']):
                    region_col = col
                    break
            
            if region_col and len(valid_sales_df) > 0:
                region_agg = valid_sales_df.groupby(region_col)[sales_col].agg(['sum', 'count', 'mean']).reset_index()
                region_agg.columns = ['Region', 'Total_Sales', 'Transaction_Count', 'Average_Sale']
                region_agg = region_agg.sort_values('Total_Sales', ascending=False)
                region_agg = region_agg.round(2)
                aggregations['regional'] = region_agg
        
        # Summary statistics
        summary_stats = {
            'total_records': len(df),
            'date_range': None,
            'total_sales': None,
            'average_sale': None
        }
        
        if 'date' in df.columns and not df['date'].isnull().all():
            summary_stats['date_range'] = f"{df['date'].min().strftime('%Y-%m-%d')} to {df['date'].max().strftime('%Y-%m-%d')}"
        
        if sales_col and len(valid_sales_df) > 0:
            # Calculate accurate statistics using only valid sales data
            summary_stats['total_sales'] = valid_sales_df[sales_col].sum()
            summary_stats['average_sale'] = valid_sales_df[sales_col].mean()
            summary_stats['valid_transactions'] = len(valid_sales_df)
            summary_stats['excluded_invalid'] = len(df) - len(valid_sales_df)
            
            if self.verbose:
                print(f"Accurate summary stats: Total=${summary_stats['total_sales']:,.0f}, Avg=${summary_stats['average_sale']:,.0f}")
                print(f"Valid transactions: {summary_stats['valid_transactions']:,}, Excluded invalid: {summary_stats['excluded_invalid']:,}")
        
        aggregations['summary'] = summary_stats
        
        if self.verbose:
            print(f"Generated {len(aggregations)} aggregation datasets")
        
        return aggregations
    
    def process_data(self, raw_data):
        """
        Complete data processing pipeline.
        
        Args:
            raw_data (pd.DataFrame): Raw input data
            
        Returns:
            dict: Processed data ready for report generation
        """
        # Step 1: Standardize columns
        standardized_data = self.standardize_columns(raw_data)
        
        # Step 2: Parse dates
        date_parsed_data = self.parse_dates(standardized_data)
        
        # Step 3: Clean data
        cleaned_data = self.clean_data(date_parsed_data)
        
        # Step 4: Generate aggregations
        aggregations = self.aggregate_data(cleaned_data)
        
        # Prepare final processed data structure
        processed_data = {
            'raw_data': cleaned_data,
            'aggregations': aggregations,
            'processing_timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        return processed_data
