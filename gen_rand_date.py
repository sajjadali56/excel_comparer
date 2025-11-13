import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

def generate_sample_excel_files():
    """Generate two sample Excel files with realistic financial data for testing"""
    
    # Set random seed for reproducibility
    np.random.seed(42)
    random.seed(42)
    
    # Sample data configurations
    companies = ['Apple Inc.', 'Microsoft Corp', 'Google LLC', 'Amazon Inc.', 'Tesla Inc.', 
                 'Meta Platforms', 'Netflix Inc.', 'NVIDIA Corp', 'Intel Corp', 'IBM Corp']
    
    products = ['Laptop', 'Smartphone', 'Tablet', 'Desktop', 'Server', 'Accessories', 'Software']
    regions = ['North America', 'Europe', 'Asia Pacific', 'Latin America', 'Middle East']
    categories = ['Technology', 'Hardware', 'Software', 'Services', 'Cloud']
    
    def create_financial_data_sheet():
        """Create Financial_Summary sheet with some differences"""
        dates = [datetime(2024, 1, 1) + timedelta(days=i*30) for i in range(12)]
        
        data_actual = {
            'Date': dates,
            'Company': [random.choice(companies) for _ in range(12)],
            'Revenue': np.random.uniform(1000000, 5000000, 12).round(2),
            'Expenses': np.random.uniform(500000, 2500000, 12).round(2),
            'Profit': np.random.uniform(200000, 1500000, 12).round(2),
            'Profit_Margin': np.random.uniform(0.05, 0.35, 12).round(3)
        }
        
        data_expected = data_actual.copy()
        # Introduce some differences
        data_expected['Revenue'] = [x * random.uniform(0.98, 1.02) for x in data_actual['Revenue']]
        data_expected['Expenses'] = [x * random.uniform(0.99, 1.03) for x in data_actual['Expenses']]
        data_expected['Profit'][3] = data_actual['Profit'][3] * 1.1  # One clear difference
        data_expected['Company'][5] = 'Microsoft Corporation'  # Name difference
        
        return pd.DataFrame(data_actual), pd.DataFrame(data_expected)
    
    def create_sales_data_sheet():
        """Create Sales_Data sheet with categorical differences"""
        data_actual = {
            'Product_ID': [f'PROD_{i:03d}' for i in range(1, 21)],
            'Product_Name': [random.choice(products) for _ in range(20)],
            'Region': [random.choice(regions) for _ in range(20)],
            'Units_Sold': np.random.randint(100, 5000, 20),
            'Unit_Price': np.random.uniform(50, 2000, 20).round(2),
            'Total_Sales': np.random.uniform(5000, 500000, 20).round(2)
        }
        
        data_expected = data_actual.copy()
        # Introduce categorical differences
        data_expected['Region'][2] = 'APAC'  # Different region name
        data_expected['Product_Name'][7] = 'Gaming Laptop'  # Different product name
        data_expected['Units_Sold'][10] = data_actual['Units_Sold'][10] + 100  # Quantity difference
        data_expected['Units_Sold'][15] = data_actual['Units_Sold'][15] - 50   # Another difference
        
        return pd.DataFrame(data_actual), pd.DataFrame(data_expected)
    
    def create_employee_data_sheet():
        """Create Employee_Info sheet with text data differences"""
        departments = ['Engineering', 'Sales', 'Marketing', 'HR', 'Finance', 'Operations']
        positions = ['Manager', 'Analyst', 'Developer', 'Director', 'Specialist', 'Coordinator']
        
        data_actual = {
            'Employee_ID': [f'EMP{i:04d}' for i in range(1, 16)],
            'Name': [f'Employee_{i}' for i in range(1, 16)],
            'Department': [random.choice(departments) for _ in range(15)],
            'Position': [random.choice(positions) for _ in range(15)],
            'Salary': np.random.uniform(50000, 150000, 15).round(2),
            'Join_Date': [datetime(2020, 1, 1) + timedelta(days=random.randint(0, 1460)) for _ in range(15)]
        }
        
        data_expected = data_actual.copy()
        # Introduce text differences
        data_expected['Department'][3] = 'Human Resources'  # Same as HR but different text
        data_expected['Name'][8] = 'Employee_008'  # Different naming convention
        data_expected['Salary'][12] = data_actual['Salary'][12] * 1.05  # Salary difference
        
        return pd.DataFrame(data_actual), pd.DataFrame(data_expected)
    
    def create_quarterly_metrics_sheet():
        """Create Quarterly_Metrics sheet with UW_Year and Loss_Period columns"""
        quarters = ['Q1-2024', 'Q2-2024', 'Q3-2024', 'Q4-2024']
        
        data_actual = {
            'UW_Year': [2024, 2024, 2024, 2024],  # Force text comparison
            'Loss_Period': ['Q1', 'Q2', 'Q3', 'Q4'],  # Force text comparison
            'Gross_Premium': np.random.uniform(1000000, 3000000, 4).round(2),
            'Claims_Paid': np.random.uniform(200000, 800000, 4).round(2),
            'Loss_Ratio': np.random.uniform(0.15, 0.35, 4).round(3),
            'Expense_Ratio': np.random.uniform(0.25, 0.45, 4).round(3)
        }
        
        data_expected = data_actual.copy()
        # Introduce differences in forced text columns
        data_expected['UW_Year'] = ['2024', '2024', '2024', '2024']  # String vs numeric
        data_expected['Loss_Period'][2] = 'Quarter3'  # Different text
        data_expected['Gross_Premium'][1] = data_actual['Gross_Premium'][1] * 0.97
        data_expected['Loss_Ratio'][3] = data_actual['Loss_Ratio'][3] + 0.02
        
        return pd.DataFrame(data_actual), pd.DataFrame(data_expected)
    
    # Generate data for all sheets
    print("Generating sample Excel files...")
    
    # File 1 - Actual Data
    with pd.ExcelWriter('Sample_Financial_Data_Actual.xlsx', engine='openpyxl') as writer:
        financial_actual, financial_expected = create_financial_data_sheet()
        sales_actual, sales_expected = create_sales_data_sheet()
        employee_actual, employee_expected = create_employee_data_sheet()
        quarterly_actual, quarterly_expected = create_quarterly_metrics_sheet()
        
        financial_actual.to_excel(writer, sheet_name='Financial_Summary', index=False)
        sales_actual.to_excel(writer, sheet_name='Sales_Data', index=False)
        employee_actual.to_excel(writer, sheet_name='Employee_Info', index=False)
        quarterly_actual.to_excel(writer, sheet_name='Quarterly_Metrics', index=False)
        
        # Add some formatting
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # File 2 - Expected Data (with intentional differences)
    with pd.ExcelWriter('Sample_Financial_Data_Expected.xlsx', engine='openpyxl') as writer:
        financial_actual, financial_expected = create_financial_data_sheet()
        sales_actual, sales_expected = create_sales_data_sheet()
        employee_actual, employee_expected = create_employee_data_sheet()
        quarterly_actual, quarterly_expected = create_quarterly_metrics_sheet()
        
        financial_expected.to_excel(writer, sheet_name='Financial_Summary', index=False)
        sales_expected.to_excel(writer, sheet_name='Sales_Data', index=False)
        employee_expected.to_excel(writer, sheet_name='Employee_Info', index=False)
        quarterly_expected.to_excel(writer, sheet_name='Quarterly_Metrics', index=False)
        
        # Add some formatting
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print("✅ Sample files generated successfully!")
    print("📁 Files created:")
    print("   - Sample_Financial_Data_Actual.xlsx")
    print("   - Sample_Financial_Data_Expected.xlsx")
    print("\n📊 File details:")
    print("   - 4 sheets each: Financial_Summary, Sales_Data, Employee_Info, Quarterly_Metrics")
    print("   - 5-6 columns per sheet with mixed data types")
    print("   - Intentional differences for testing comparison tool")
    print("\n🎯 Test scenarios included:")
    print("   - Numeric differences in financial metrics")
    print("   - Text/categorical differences in product/region names")
    print("   - UW_Year and Loss_Period columns (forced text comparison)")
    print("   - Mixed data types and formatting")

def generate_detailed_readme():
    """Generate a detailed README about the sample files"""
    readme_content = """
# Sample Excel Files for ExcelCompare Pro Testing

## Overview
These sample Excel files are designed to test the ExcelCompare Pro tool with realistic financial data. The files contain intentional differences to demonstrate the tool's comparison capabilities.

## File Structure

### Sample_Financial_Data_Actual.xlsx
- **Financial_Summary**: Monthly financial data with revenue, expenses, profit metrics
- **Sales_Data**: Product sales information across different regions
- **Employee_Info**: Employee details with salaries and department information
- **Quarterly_Metrics**: Insurance/financial metrics with UW_Year and Loss_Period columns

### Sample_Financial_Data_Expected.xlsx
Same structure as Actual file, but with intentional differences for testing.

## Intentional Differences Included

### 1. Financial_Summary Sheet
- Revenue and expense values differ by ±2-3%
- One clear profit difference (10% variation)
- Company name variation (Microsoft Corp vs Microsoft Corporation)

### 2. Sales_Data Sheet
- Region name differences (Asia Pacific vs APAC)
- Product name variations
- Units sold differences (+100, -50 in specific rows)

### 3. Employee_Info Sheet
- Department name variations (HR vs Human Resources)
- Employee naming convention differences
- Salary variations

### 4. Quarterly_Metrics Sheet
- UW_Year as numbers vs strings (2024 vs "2024")
- Loss_Period text variations (Q3 vs Quarter3)
- Financial metric differences

## How to Use with ExcelCompare Pro

1. Upload `Sample_Financial_Data_Actual.xlsx` as "Actual File"
2. Upload `Sample_Financial_Data_Expected.xlsx` as "Expected File"
3. Run comparison
4. Observe how the tool detects:
   - Numeric statistical differences
   - Text value count differences
   - Mixed data type handling
   - Special column handling (UW_Year, Loss_Period)

## Data Characteristics

- **Mixed Data Types**: Numeric, text, dates
- **Realistic Ranges**: Financial data in appropriate scales
- **Common Differences**: Typical variations found in real-world scenarios
- **Special Columns**: UW_Year and Loss_Period for forced text comparison testing

These files are completely synthetic and contain no real financial data.
"""
    
    with open('SAMPLE_FILES_README.md', 'w') as f:
        f.write(readme_content)
    print("📖 Detailed README generated: SAMPLE_FILES_README.md")

if __name__ == "__main__":
    generate_sample_excel_files()
    generate_detailed_readme()
    
    print("\n" + "="*60)
    print("🎉 Sample generation complete!")
    print("="*60)
    print("\nNext steps:")
    print("1. Run the ExcelCompare Pro tool")
    print("2. Upload these sample files for testing")
    print("3. Observe the comparison results across different data types")
    print("4. Check the PDF report generation")
    print("\nThe files demonstrate:")
    print("✅ Numeric column statistical comparisons")
    print("✅ Text column value count differences") 
    print("✅ Mixed data type handling")
    print("✅ Special column processing (UW_Year, Loss_Period)")
    print("✅ Realistic financial data scenarios")
    