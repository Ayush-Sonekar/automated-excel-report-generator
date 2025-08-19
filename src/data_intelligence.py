"""
Data Intelligence Module
Smart detection of data types and automatic business intelligence insights.
"""

import pandas as pd
import numpy as np
from collections import Counter
import re

class DataIntelligence:
    """Analyzes datasets to provide intelligent insights and recommendations."""
    
    def __init__(self, verbose=False):
        self.verbose = verbose
        
    def detect_data_type(self, df):
        """
        Detect the primary business type of the dataset.
        
        Args:
            df (pd.DataFrame): Input dataset
            
        Returns:
            dict: Data type classification and confidence
        """
        column_names = [col.lower() for col in df.columns]
        column_text = ' '.join(column_names)
        
        # Define business domain keywords
        domains = {
            'sales': ['sales', 'revenue', 'amount', 'price', 'product', 'customer', 'order', 'transaction'],
            'real_estate': ['property', 'address', 'building', 'lot', 'square', 'feet', 'borough', 'neighborhood'],
            'financial': ['balance', 'account', 'credit', 'debit', 'investment', 'portfolio', 'stock', 'bond'],
            'hr': ['employee', 'salary', 'department', 'hire', 'performance', 'review'],
            'marketing': ['campaign', 'click', 'impression', 'conversion', 'lead', 'funnel'],
            'inventory': ['stock', 'warehouse', 'supplier', 'quantity', 'sku', 'category']
        }
        
        scores = {}
        for domain, keywords in domains.items():
            score = sum(1 for keyword in keywords if keyword in column_text)
            scores[domain] = score
        
        # Find best match
        if scores:
            best_domain = max(scores.keys(), key=lambda x: scores[x])
            confidence = scores[best_domain] / len(domains[best_domain]) * 100
        else:
            best_domain = 'general'
            confidence = 0
        
        return {
            'domain': best_domain,
            'confidence': min(confidence, 100),
            'all_scores': scores
        }
    
    def generate_insights(self, df, aggregations, domain_info):
        """
        Generate intelligent business insights based on data analysis.
        
        Args:
            df (pd.DataFrame): Processed dataset
            aggregations (dict): Data aggregations
            domain_info (dict): Domain detection results
            
        Returns:
            list: List of business insights
        """
        insights = []
        domain = domain_info['domain']
        
        # Domain-specific insights
        if domain == 'sales':
            insights.extend(self._generate_sales_insights(df, aggregations))
        elif domain == 'real_estate':
            insights.extend(self._generate_real_estate_insights(df, aggregations))
        elif domain == 'financial':
            insights.extend(self._generate_financial_insights(df, aggregations))
        
        # General statistical insights
        insights.extend(self._generate_statistical_insights(df))
        
        return insights
    
    def _generate_sales_insights(self, df, aggregations):
        """Generate sales-specific insights."""
        insights = []
        
        # Monthly trends
        if 'monthly' in aggregations:
            monthly = aggregations['monthly']
            if not monthly.empty and len(monthly) > 1:
                best_month = monthly.loc[monthly.iloc[:, 1].idxmax(), monthly.columns[0]]
                worst_month = monthly.loc[monthly.iloc[:, 1].idxmin(), monthly.columns[0]]
                insights.append(f"Peak sales month: {best_month}")
                insights.append(f"Lowest sales month: {worst_month}")
                
                # Growth trend
                if len(monthly) >= 3:
                    recent_avg = monthly.iloc[-3:, 1].mean()
                    earlier_avg = monthly.iloc[:3, 1].mean()
                    if recent_avg > earlier_avg * 1.1:
                        insights.append("Sales show strong growth trend in recent months")
                    elif recent_avg < earlier_avg * 0.9:
                        insights.append("Sales show declining trend - attention needed")
        
        # Product analysis
        if 'product' in aggregations:
            products = aggregations['product']
            if not products.empty:
                top_product = products.iloc[0, 0]
                top_revenue = products.iloc[0, 1]
                insights.append(f"Star product '{top_product}' generates ${top_revenue:,.2f}")
                
                if len(products) > 5:
                    top_5_revenue = products.iloc[:5, 1].sum()
                    total_revenue = products.iloc[:, 1].sum()
                    concentration = (top_5_revenue / total_revenue) * 100
                    insights.append(f"Top 5 products account for {concentration:.1f}% of total revenue")
        
        return insights
    
    def _generate_real_estate_insights(self, df, aggregations):
        """Generate real estate-specific insights."""
        insights = []
        
        # Price analysis
        if 'sale_price' in df.columns:
            # Clean price data - convert to numeric, handling strings and invalid values
            prices = pd.to_numeric(df['sale_price'].astype(str).str.replace(',', '').str.replace(' ', '').str.replace('-', '0'), errors='coerce')
            prices = prices.replace(0, np.nan).dropna()
            
            if not prices.empty and len(prices) > 0:
                median_price = prices.median()
                mean_price = prices.mean()
                insights.append(f"Median property price: ${median_price:,.0f}")
                
                if mean_price > median_price * 1.3:
                    insights.append("High-end properties significantly skew average prices")
        
        # Geographic insights
        if 'regional' in aggregations:
            regions = aggregations['regional']
            if not regions.empty:
                top_region = regions.iloc[0, 0]
                if str(top_region) != 'Unknown' and str(top_region) != 'nan':
                    insights.append(f"Most active market area: {top_region}")
        
        # Building type analysis
        building_cols = [col for col in df.columns if 'building' in col.lower() and 'class' in col.lower()]
        if building_cols:
            building_col = building_cols[0]
            building_types = df[building_col].value_counts()
            if not building_types.empty:
                top_type = building_types.index[0]
                if str(top_type) != 'Unknown' and str(top_type) != 'nan':
                    insights.append(f"Most common property type: {top_type}")
        
        # Neighborhood analysis
        neighborhood_cols = [col for col in df.columns if 'neighborhood' in col.lower() or 'area' in col.lower()]
        if neighborhood_cols:
            neighborhood_col = neighborhood_cols[0]
            neighborhoods = df[neighborhood_col].value_counts().head(3)
            if not neighborhoods.empty:
                top_neighborhood = neighborhoods.index[0]
                if str(top_neighborhood) != 'Unknown':
                    insights.append(f"Most active neighborhood: {top_neighborhood}")
        
        # Borough analysis for NYC data
        if 'borough' in [col.lower() for col in df.columns]:
            borough_col = next(col for col in df.columns if col.lower() == 'borough')
            boroughs = df[borough_col].value_counts()
            if not boroughs.empty:
                top_borough = boroughs.index[0]
                borough_names = {1: 'Manhattan', 2: 'Bronx', 3: 'Brooklyn', 4: 'Queens', 5: 'Staten Island'}
                borough_name = borough_names.get(top_borough, top_borough)
                insights.append(f"Most active borough: {borough_name}")
        
        # Transaction volume insights
        total_transactions = len(df)
        if total_transactions > 0:
            if 'sale_price' in df.columns:
                valid_prices = df['sale_price'].dropna()
                valid_transactions = len(valid_prices)
                if valid_transactions > 0:
                    insights.append(f"{valid_transactions:,} valid property transactions analyzed")
                    if valid_transactions < total_transactions * 0.5:
                        insights.append("Note: Many transactions have missing price data")
        
        return insights
    
    def _generate_financial_insights(self, df, aggregations):
        """Generate financial-specific insights."""
        insights = []
        
        # Add financial-specific logic here
        insights.append("Financial data analysis completed")
        
        return insights
    
    def _generate_statistical_insights(self, df):
        """Generate general statistical insights."""
        insights = []
        
        # Data quality insights
        total_records = len(df)
        missing_data = df.isnull().sum().sum()
        missing_percent = (missing_data / (total_records * len(df.columns))) * 100
        
        if missing_percent < 5:
            insights.append("Excellent data quality - minimal missing values")
        elif missing_percent > 20:
            insights.append("Data quality concern - significant missing values detected")
        
        # Numeric column analysis
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols[:3]:  # Analyze first 3 numeric columns
            if col in df.columns:
                values = df[col].dropna()
                if not values.empty:
                    skewness = values.skew()
                    if abs(skewness) > 2:
                        distribution = "heavily skewed" if abs(skewness) > 3 else "moderately skewed"
                        insights.append(f"{col.title()} data is {distribution} - consider log transformation")
        
        return insights
    
    def analyze_correlations(self, df):
        """Find interesting correlations in the dataset."""
        numeric_df = df.select_dtypes(include=[np.number])
        
        if len(numeric_df.columns) < 2:
            return []
        
        correlations = []
        corr_matrix = numeric_df.corr()
        
        # Find strong correlations (excluding self-correlations)
        for i in range(len(corr_matrix.columns)):
            for j in range(i+1, len(corr_matrix.columns)):
                corr_val = corr_matrix.iloc[i, j]
                if abs(corr_val) > 0.7:  # Strong correlation threshold
                    col1 = corr_matrix.columns[i]
                    col2 = corr_matrix.columns[j]
                    direction = "positively" if corr_val > 0 else "negatively"
                    correlations.append({
                        'column1': col1,
                        'column2': col2,
                        'correlation': corr_val,
                        'description': f"{col1} and {col2} are strongly {direction} correlated ({corr_val:.3f})"
                    })
        
        return correlations
    
    def detect_outliers(self, df, columns=None):
        """Detect outliers using IQR method."""
        if columns is None:
            columns = df.select_dtypes(include=[np.number]).columns
        
        outlier_info = {}
        
        for col in columns:
            if col in df.columns:
                Q1 = df[col].quantile(0.25)
                Q3 = df[col].quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                
                outliers = df[(df[col] < lower_bound) | (df[col] > upper_bound)]
                outlier_info[col] = {
                    'count': len(outliers),
                    'percentage': (len(outliers) / len(df)) * 100,
                    'bounds': (lower_bound, upper_bound)
                }
        
        return outlier_info