"""
Configuration Module
Contains configuration settings and constants for the report generator.
"""

import os
from datetime import datetime

class Config:
    """Configuration class containing all application settings."""
    
    def __init__(self):
        # File paths and naming
        self.DEFAULT_INPUT_FILE = "sample_data.csv"
        self.OUTPUT_PREFIX = "Monthly_Sales_Report"
        self.CHART_DIR = "charts"
        
        # Data processing settings
        self.MAX_PRODUCT_DISPLAY = 20
        self.DATE_FORMAT = "%Y-%m-%d"
        self.DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"
        
        # Excel formatting settings
        self.HEADER_COLOR = "2E86AB"
        self.FONT_FAMILY = "Arial"
        self.HEADER_FONT_SIZE = 12
        self.DATA_FONT_SIZE = 10
        self.TITLE_FONT_SIZE = 16
        
        # Chart settings
        self.CHART_WIDTH = 600
        self.CHART_HEIGHT = 400
        self.CHART_DPI = 300
        self.CHART_COLORS = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#593E2C', '#8338EC']
        
        # Data validation settings
        self.MIN_RECORDS_FOR_ANALYSIS = 1
        self.MAX_MISSING_DATA_PERCENT = 50
        
    def get_timestamped_filename(self, base_name=None, extension=".xlsx"):
        """
        Generate a timestamped filename.
        
        Args:
            base_name (str): Base name for the file
            extension (str): File extension
            
        Returns:
            str: Timestamped filename
        """
        if base_name is None:
            base_name = self.OUTPUT_PREFIX
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base_name}_{timestamp}{extension}"
    
    def get_chart_path(self, chart_name):
        """
        Get the path for a chart file.
        
        Args:
            chart_name (str): Name of the chart
            
        Returns:
            str: Full path to chart file
        """
        return os.path.join(self.CHART_DIR, f"{chart_name}.png")
    
    def ensure_chart_directory(self):
        """Ensure chart directory exists."""
        os.makedirs(self.CHART_DIR, exist_ok=True)
    
    @property
    def supported_date_formats(self):
        """List of supported date formats for parsing."""
        return [
            "%Y-%m-%d",
            "%m/%d/%Y",
            "%d/%m/%Y",
            "%Y/%m/%d",
            "%m-%d-%Y",
            "%d-%m-%Y",
            "%Y%m%d",
            "%m/%d/%y",
            "%d/%m/%y"
        ]
    
    @property
    def numeric_columns_keywords(self):
        """Keywords that indicate numeric/currency columns."""
        return [
            'sales', 'revenue', 'amount', 'total', 'value', 'price', 
            'cost', 'profit', 'margin', 'quantity', 'units', 'count'
        ]
    
    @property
    def date_columns_keywords(self):
        """Keywords that indicate date columns."""
        return [
            'date', 'time', 'created', 'updated', 'order', 'transaction',
            'purchase', 'sale', 'timestamp'
        ]
    
    @property
    def category_columns_keywords(self):
        """Keywords that indicate categorical columns."""
        return [
            'product', 'item', 'category', 'type', 'region', 'location',
            'area', 'territory', 'customer', 'client', 'segment'
        ]
