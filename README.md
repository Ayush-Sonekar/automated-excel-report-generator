# Automated Excel Report Generator

An intelligent Python application that processes business data from CSV files and generates professional Excel reports with charts, formatting, and **smart domain-specific insights**.

## 🚀 New Smart Features

### Smart Data Detection & Templates
The system now automatically detects your data type and provides tailored business insights:

- **Sales Data**: Product performance, monthly trends, regional analysis
- **Real Estate**: Property price analysis, market insights, geographic trends
- **Financial**: Account analysis, investment tracking, portfolio insights
- **HR Data**: Employee metrics, department analysis, performance tracking
- **Marketing**: Campaign analysis, conversion tracking, lead funnel insights
- **Inventory**: Stock analysis, supplier performance, warehouse optimization

## 📊 What It Does

1. **Automatically detects** your business domain with confidence scoring
2. **Generates domain-specific insights** tailored to your data type
3. **Creates professional Excel reports** with multiple analysis sheets
4. **Embeds interactive charts** and visualizations
5. **Provides executive summaries** with intelligent business insights

## 🔧 Quick Start

### Basic Usage
```bash
# Process any CSV file - the system automatically detects the data type
python main.py your_data.csv

# Enable detailed logging
python main.py your_data.csv --verbose

# Specify custom output name
python main.py your_data.csv --output "My_Report.xlsx"
```

### Example Output
```
Processing data from: sample_data.csv
Analyzing data type and generating insights...
Detected data type: sales (confidence: 85.0%)
Processing and aggregating data...
Report successfully generated: Monthly_Sales_Report_20250813_065337.xlsx
```

## 📈 Generated Reports Include

- **Executive Summary** with smart insights based on detected domain
- **Raw Data** with professional formatting
- **Monthly Analysis** with trend charts
- **Product/Item Performance** rankings
- **Regional Distribution** analysis  
- **Interactive Charts** embedded in Excel

## 🎯 Smart Insights Examples

### For Sales Data:
- "Star product 'Laptop Pro' generates $12,000.00"
- "Peak sales month: March"
- "Top 5 products account for 78.2% of total revenue"

### For Real Estate Data:
- "Median property price: $485,000"
- "Most active market area: Manhattan"
- "High-end properties significantly skew average prices"

## 📋 Requirements

- Python 3.7+
- pandas, openpyxl, matplotlib, numpy (automatically installed)

## 🔧 Installation

```bash
# All dependencies are automatically managed
python main.py your_file.csv
```

## 💡 Supported Data Types

The system works with any CSV file containing business data. It automatically detects and optimizes for:

- Sales transactions
- Real estate listings
- Financial records
- Employee data
- Marketing campaigns
- Inventory management
- Customer data
- Any structured business data

## 🎨 Features

- **Smart Domain Detection**: Automatically identifies your business data type
- **Professional Formatting**: Corporate-ready Excel reports
- **Interactive Charts**: Multiple visualization types embedded in Excel
- **Multi-encoding Support**: Handles various CSV formats and encodings
- **Robust Data Cleaning**: Handles missing values and malformed data
- **Command Line Interface**: Easy automation and scripting
- **Verbose Logging**: Detailed processing information

## 📊 Report Structure

1. **Executive Summary**: Smart insights based on detected domain
2. **Raw Data**: Cleaned and formatted source data
3. **Monthly Analysis**: Time-based trends and patterns
4. **Product/Category Analysis**: Performance rankings
5. **Regional Analysis**: Geographic distribution
6. **Charts**: Visual summaries and dashboards

---

*Built with intelligent data analysis capabilities - transforms any business CSV into professional reports with domain-specific insights.*
