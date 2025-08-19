"""
Chart Generation Module
Creates matplotlib charts and handles chart embedding into Excel files.
"""

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
import numpy as np
from io import BytesIO
import base64
from datetime import datetime

class ChartGenerator:
    """Generates various types of charts for Excel reports."""
    
    def __init__(self, style='seaborn-v0_8', figsize=(12, 8)):
        """
        Initialize chart generator with styling options.
        
        Args:
            style (str): Matplotlib style to use
            figsize (tuple): Default figure size
        """
        self.figsize = figsize
        # Set style if available, otherwise use default
        try:
            plt.style.use(style)
        except:
            plt.style.use('default')
        
        # Set color palette
        self.colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#593E2C', '#8338EC']
        
    def create_monthly_trend_chart(self, monthly_data, save_path=None):
        """
        Create a monthly sales trend line chart.
        
        Args:
            monthly_data (pd.DataFrame): Monthly aggregated data
            save_path (str): Path to save chart image
            
        Returns:
            str: Path to saved chart or None
        """
        if monthly_data.empty:
            return None
            
        fig, ax = plt.subplots(figsize=self.figsize)
        
        # Extract data
        months = monthly_data.iloc[:, 0]  # First column should be months
        sales = monthly_data.iloc[:, 1]   # Second column should be total sales
        
        # Create line plot
        ax.plot(months, sales, marker='o', linewidth=3, markersize=8, 
                color=self.colors[0], markerfacecolor=self.colors[1])
        
        # Formatting
        ax.set_title('Monthly Sales Trend', fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('Month', fontsize=12, fontweight='bold')
        ax.set_ylabel('Total Sales', fontsize=12, fontweight='bold')
        
        # Rotate x-axis labels for better readability
        plt.xticks(rotation=45, ha='right')
        
        # Add grid
        ax.grid(True, alpha=0.3)
        
        # Format y-axis as currency if values are large
        if sales.max() > 1000:
            from matplotlib.ticker import FuncFormatter
            ax.yaxis.set_major_formatter(FuncFormatter(lambda x, p: f'${x:,.0f}'))
        
        # Tight layout to prevent label cutoff
        plt.tight_layout()
        
        # Save chart
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            plt.close()
            return save_path
        else:
            # Return as bytes for embedding
            buffer = BytesIO()
            plt.savefig(buffer, format='png', dpi=300, bbox_inches='tight')
            buffer.seek(0)
            plt.close()
            return buffer
    
    def create_product_performance_chart(self, product_data, save_path=None, top_n=10):
        """
        Create a horizontal bar chart for top product performance.
        
        Args:
            product_data (pd.DataFrame): Product aggregated data
            save_path (str): Path to save chart image
            top_n (int): Number of top products to show
            
        Returns:
            str: Path to saved chart or None
        """
        if product_data.empty:
            return None
            
        # Get top N products
        top_products = product_data.head(top_n)
        
        fig, ax = plt.subplots(figsize=(12, max(8, top_n * 0.6)))
        
        # Extract data
        products = top_products.iloc[:, 0]  # First column should be products
        sales = top_products.iloc[:, 1]     # Second column should be sales
        
        # Create horizontal bar chart
        bars = ax.barh(range(len(products)), sales, color=self.colors[0], alpha=0.8)
        
        # Formatting
        ax.set_title(f'Top {len(products)} Products by Sales', fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('Total Sales', fontsize=12, fontweight='bold')
        ax.set_ylabel('Products', fontsize=12, fontweight='bold')
        
        # Set product names as y-tick labels
        ax.set_yticks(range(len(products)))
        ax.set_yticklabels(products, fontsize=10)
        
        # Add value labels on bars
        for i, bar in enumerate(bars):
            width = bar.get_width()
            ax.text(width + width*0.01, bar.get_y() + bar.get_height()/2, 
                   f'${width:,.0f}' if width > 1000 else f'{width:.1f}',
                   ha='left', va='center', fontsize=9)
        
        # Format x-axis as currency if values are large
        if sales.max() > 1000:
            from matplotlib.ticker import FuncFormatter
            ax.xaxis.set_major_formatter(FuncFormatter(lambda x, p: f'${x:,.0f}'))
        
        # Add grid
        ax.grid(True, alpha=0.3, axis='x')
        
        # Invert y-axis to show highest values at top
        ax.invert_yaxis()
        
        plt.tight_layout()
        
        # Save chart
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            plt.close()
            return save_path
        else:
            buffer = BytesIO()
            plt.savefig(buffer, format='png', dpi=300, bbox_inches='tight')
            buffer.seek(0)
            plt.close()
            return buffer
    
    def create_regional_pie_chart(self, regional_data, save_path=None):
        """
        Create a pie chart for regional sales distribution.
        
        Args:
            regional_data (pd.DataFrame): Regional aggregated data
            save_path (str): Path to save chart image
            
        Returns:
            str: Path to saved chart or None
        """
        if regional_data.empty:
            return None
            
        fig, ax = plt.subplots(figsize=(10, 10))
        
        # Extract data
        regions = regional_data.iloc[:, 0]  # First column should be regions
        sales = regional_data.iloc[:, 1]    # Second column should be sales
        
        # Create pie chart
        pie_result = ax.pie(sales, labels=regions, autopct='%1.1f%%',
                           startangle=90, colors=self.colors[:len(regions)])
        
        # Handle different return formats from pie chart
        if len(pie_result) == 3:
            wedges, texts, autotexts = pie_result
        else:
            wedges, texts = pie_result
            autotexts = []
        
        # Formatting
        ax.set_title('Sales Distribution by Region', fontsize=16, fontweight='bold', pad=20)
        
        # Make percentage text bold and larger
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(10)
        
        # Equal aspect ratio ensures that pie is drawn as a circle
        ax.axis('equal')
        
        plt.tight_layout()
        
        # Save chart
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            plt.close()
            return save_path
        else:
            buffer = BytesIO()
            plt.savefig(buffer, format='png', dpi=300, bbox_inches='tight')
            buffer.seek(0)
            plt.close()
            return buffer
    
    def create_summary_dashboard(self, summary_stats, save_path=None):
        """
        Create a summary dashboard with key metrics.
        
        Args:
            summary_stats (dict): Summary statistics
            save_path (str): Path to save chart image
            
        Returns:
            str: Path to saved chart or None
        """
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
        
        # Remove axes for text-based dashboard
        for ax in [ax1, ax2, ax3, ax4]:
            ax.axis('off')
        
        # Metric 1: Total Records
        ax1.text(0.5, 0.5, f"{summary_stats.get('total_records', 'N/A'):,}", 
                ha='center', va='center', fontsize=48, fontweight='bold', 
                color=self.colors[0], transform=ax1.transAxes)
        ax1.text(0.5, 0.2, 'Total Records', ha='center', va='center', 
                fontsize=16, transform=ax1.transAxes)
        
        # Metric 2: Total Sales
        total_sales = summary_stats.get('total_sales', 0)
        sales_text = f"${total_sales:,.0f}" if total_sales and total_sales > 1000 else str(total_sales or 'N/A')
        ax2.text(0.5, 0.5, sales_text, ha='center', va='center', 
                fontsize=48, fontweight='bold', color=self.colors[1], 
                transform=ax2.transAxes)
        ax2.text(0.5, 0.2, 'Total Sales', ha='center', va='center', 
                fontsize=16, transform=ax2.transAxes)
        
        # Metric 3: Average Sale
        avg_sale = summary_stats.get('average_sale', 0)
        avg_text = f"${avg_sale:.2f}" if avg_sale else 'N/A'
        ax3.text(0.5, 0.5, avg_text, ha='center', va='center', 
                fontsize=48, fontweight='bold', color=self.colors[2], 
                transform=ax3.transAxes)
        ax3.text(0.5, 0.2, 'Average Sale', ha='center', va='center', 
                fontsize=16, transform=ax3.transAxes)
        
        # Metric 4: Date Range
        date_range = summary_stats.get('date_range', 'N/A')
        ax4.text(0.5, 0.5, date_range, ha='center', va='center', 
                fontsize=24, fontweight='bold', color=self.colors[3], 
                transform=ax4.transAxes)
        ax4.text(0.5, 0.2, 'Data Period', ha='center', va='center', 
                fontsize=16, transform=ax4.transAxes)
        
        # Add borders around each metric
        from matplotlib.patches import Rectangle
        for ax in [ax1, ax2, ax3, ax4]:
            rect = Rectangle((0.05, 0.05), 0.9, 0.9, linewidth=2, 
                           edgecolor=self.colors[0], facecolor='none', 
                           transform=ax.transAxes)
            ax.add_patch(rect)
        
        plt.suptitle('Sales Summary Dashboard', fontsize=20, fontweight='bold', y=0.95)
        plt.tight_layout()
        
        # Save chart
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            plt.close()
            return save_path
        else:
            buffer = BytesIO()
            plt.savefig(buffer, format='png', dpi=300, bbox_inches='tight')
            buffer.seek(0)
            plt.close()
            return buffer
    
    def generate_all_charts(self, aggregations, output_dir='.'):
        """
        Generate all available charts based on aggregation data.
        
        Args:
            aggregations (dict): All aggregated datasets
            output_dir (str): Directory to save charts
            
        Returns:
            dict: Dictionary of chart paths/buffers
        """
        charts = {}
        
        # Monthly trend chart
        if 'monthly' in aggregations and not aggregations['monthly'].empty:
            charts['monthly_trend'] = self.create_monthly_trend_chart(aggregations['monthly'])
        
        # Product performance chart
        if 'product' in aggregations and not aggregations['product'].empty:
            charts['product_performance'] = self.create_product_performance_chart(aggregations['product'])
        
        # Regional pie chart
        if 'regional' in aggregations and not aggregations['regional'].empty:
            charts['regional_distribution'] = self.create_regional_pie_chart(aggregations['regional'])
        
        # Summary dashboard
        if 'summary' in aggregations:
            charts['summary_dashboard'] = self.create_summary_dashboard(aggregations['summary'])
        
        return charts
