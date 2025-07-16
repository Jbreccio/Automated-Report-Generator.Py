# 📊 Automated Report Generator

A complete system for automatically generating Excel reports with advanced analysis, graphs, and professional formatting. Transforms raw data into presentation-ready executive reports.

## 🚀 Features

- **📈 Automatic Report Generation**: Creates complete Excel reports automatically
- **🎨 Professional Formatting**: Applies styles, colors, and automatic formatting
- **📊 Integrated Charts**: Automatically generates bar and line charts
- **🔍 Advanced Analytics**: Calculates KPIs, trends, and statistics
- **📋 Multiple Worksheets**: Organizes data into themed worksheets
- **💼 Executive Summary**: Creates a summary spreadsheet with key insights
- **🎯 Top Performers**: Identify top sellers, products, and regions
- **📅 Time-Lapse Analysis**: Tracks trends over time
- **🔄 Simulated Data**: Generates sample data for demonstration purposes

## 🛠️ Technologies Used

- **Python 3.7+**
- **pandas**: Data manipulation and analysis Data
- **openpyxl**: Creating and formatting Excel files
- **numpy**: Numerical and statistical calculations
- **matplotlib**: Basic visualizations
- **seaborn**: Advanced statistical graphics
- **dataclasses**: Typed data structures

## 📋 Prerequisites

```bash
Python 3.7 or higher
pip (Python package manager)
```

## 🔧 Installation

1. **Clone the repository:**
```bash
git clone https://github.com/your-user/automated-report-generator.git
cd automated-report-generator
```

2. **Install the dependencies:**
```bash
pip install -r requirements.txt
```

Or install manually:
```bash
pip install pandas openpyxl numpy matplotlib seaborn
```

## 🚀 How to Use

### Basic Execution

```bash
python main.py
```

### Usage Examples

#### 1. Generate Sales Report
```python
from report_generator import ExcelReportGenerator, ReportConfig, DataGenerator

# Configuration
config = ReportConfig(
title="Sales Report Q1",
output_path="relatorio_vendas_q1.xlsx",
company_name="My Company"
)

# Generate Data
sales_data = DataGenerator.generate_sales_data(90)

# Create Report
generator = ExcelReportGenerator(config)
generator.generate_complete_report({"Sales": sales_data})
```

#### 2. Financial Report
```python
# Generate Financial Data
financial_data = DataGenerator.generate_financial_data(12)

# Configure Report
config = ReportConfig(
title="Annual Financial Report",
output_path="relatorio_financeiro_2024.xlsx",
include_charts=True,
include_summary=True
)

# Generate Report
data_dict = {
"Financial Data": financial_data,
"Monthly Analysis": financial_data.groupby('mes').sum()
}

generator = ExcelReportGenerator(config)
generator.generate_complete_report(data_dict)
```

#### 3. Custom Analysis
```python
from report_generator import ReportAnalyzer

# Analyze Data
analyzer = ReportAnalyzer(sales_data)

# Summary Statistics
stats = analyzer.get_summary_stats()
print(f"Total records: {stats['total_records']}")

# Top performers
top_sellers = analyzer.get_top_performers('seller', 'net_value', 5)
top_products = analyzer.get_top_performers('product', 'net_value', 5)

# Trend analysis
trend = analyzer.get_trend_analysis('data', 'net_value')
```

## ⚙️ Configuration

### ReportConfig Parameters

```python
config = ReportConfig(
title="Report Title", # Main Title
output_path="path/file.xlsx", # Output Path
include_charts=True, # Include charts
include_summary=True, # Include executive summary
auto_format=True, # Automatic formatting
company_name="Company Name" # Company name
)
```

## 📁 Project Structure

```
automated-report-generator/
├── main.py # Main file
├── report_generator.py # Generator classes
├── requirements.txt # Dependencies
├── README.md # Documentation
├── reports/ # Generated reports
│ ├── sales_report.xlsx
│ ├── financial_report.xlsx
│ └── executive_report.xlsx
├── data/ # Input data
│ ├── sales_data.csv
│ └── financial_data.csv
└── templates/ # Report Templates
└── template_padrao.xlsx
```

## 📊 Generated Report Types

### 1. Sales Report
- **Detailed Sales**: All sales records
- **Top Sellers**: Ranking by performance
- **Top Products**: Best-selling products
- **Analysis by Region**: Geographic distribution
- **Monthly Trends**: Evolution over time

### 2. Financial Report
- **Financial Data**: Revenue, costs, and profits
- **Margin Analysis**: Gross and net margin
- **Monthly Growth**: Growth rate
- **Cost Analysis**: Distribution of Expenses
- **Projections**: Future Trends

### 3. Executive Summary

Main KPIs: Key metrics
Overall Statistics: Data summary
Insights: Points of attention
Recommendations: Data-based suggestions

🎨 Automatic Formatting
Applied Styles

Headings: Blue background, white text, bold
Borders: Thin borders on all cells
Alignment: Centered for headers
Width: Auto-adjust based on content
Colors: Professional color scheme

Automatic Charts

Bar Charts: For comparisons
Line Charts: For time trends
Positioning: Automatic next to data
Styles: Predefined and professional

📈 Available Analyses
Basic Statistics
pythonsummary = analyzer.get_summary_stats()
# Returns: total_records, date_range, numeric_summary
Top Performers
pythontop_items = analyzer.get_top_performers('column', 'value', 10)
# Returns: ranking ordered by value
Trend Analysis
pythontrend = analyzer.get_trend_analysis('data', 'value')
# Returns: data aggregated by period with growth rate
🔍 Example of Generated Data
Sales Data
python{
'date': '2024-01-15',
'product': 'Product A',
'salesperson': 'John Smith',
'quantity': 5,
'unit_price': 299.99,
'net_value': 1274.95,
'region': 'South'
}
Financial Data
python{
'month': '2024-01',
'revenue': 150000.00,
'costs': 90000.00,
'net_profit': 45000.00,
'net_margin': 30.00
}
🎯 Use Cases
For Businesses

Monthly Reports: Automation of recurring reports
Performance Analysis: KPI tracking
Presentations: Reports for executive meetings

For Analysts

Data Analysis: Quickly transform data into insights
Visualizations: Automated professional charts
Ad-hoc Reports: Custom analyses

For Developers

Integration: Easy integration with existing systems
Customization: Modular and extensible code
Automation: Programmatic report generation

🤝 Contributing

Fork the project
Create a branch for your feature (git checkout -b feature/NewFunction)
Commit your changes (git commit -m 'Add new feature')
Push to the branch (git push origin feature/NewFunction)
Open a Pull Request

📝 Upcoming Features

Web Dashboard: Web interface for configuration
Multiple Formats: PDF, PowerPoint support
Data Connectors: Database integration
Custom Templates: Template system
Scheduling: Scheduled automatic execution
Notifications: Automatic email sending
APIs: REST endpoints for integration
Interactive Reports: Dynamic Dashboards

🐛 Troubleshooting
Error: "ModuleNotFoundError"
bashpip install -r requirements.txt
Error: "Permission denied"
bash# Make sure the Excel file is not open
# Or change the output path
Error: "Invalid data format"
bash# Make sure the data is in pandas DataFrame format
# Use DataGenerator for generating sample data

📄 License
This project is licensed under the MIT License. See the LICENSE file for more details.
👨‍💻 Author
Your Name

GitHub: @Jbreccio
LinkedIn: www.linkedin.com/in/josebreccio-dev-35b8292a4
Email: oibreccio@hotmail.com

🙏 Thanks

Python community for the excellent documentation
Openpyxl library developers
Pandas community for making data analysis easier
Everyone who contributed feedback and suggestions

⭐ If this project helped you, leave a star! ⭐
