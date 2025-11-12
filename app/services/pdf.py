from fpdf import FPDF
import json
from datetime import datetime
import os

class PDFReport(FPDF):
    def __init__(self, comparison_data):
        super().__init__()
        self.comparison_data = comparison_data
        self.set_auto_page_break(auto=True, margin=15)
        
    def header(self):
        # Logo and header
        self.set_font('Arial', 'B', 16)
        self.set_text_color(31, 73, 125)  # Dark blue
        self.cell(0, 10, 'Excel Comparison Report', 0, 1, 'C')
        self.set_font('Arial', 'I', 10)
        self.set_text_color(100, 100, 100)
        self.cell(0, 5, f'Generated on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 0, 1, 'C')
        self.ln(5)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')
    
    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.set_fill_color(240, 240, 240)
        self.set_text_color(31, 73, 125)
        self.cell(0, 10, title, 0, 1, 'L', 1)
        self.ln(4)
    
    def summary_section(self):
        self.chapter_title('COMPARISON SUMMARY')
        
        # File information
        self.set_font('Arial', 'B', 10)
        self.cell(40, 8, 'Files Compared:', 0, 0)
        self.set_font('Arial', '', 10)
        self.cell(0, 8, f"{self.comparison_data['file1_name']} vs {self.comparison_data['file2_name']}", 0, 1)
        
        self.set_font('Arial', 'B', 10)
        self.cell(40, 8, 'Comparison Time:', 0, 0)
        self.set_font('Arial', '', 10)
        self.cell(0, 8, self.comparison_data['comparison_time'], 0, 1)
        
        self.ln(5)
        
        # Summary statistics
        if 'summary' in self.comparison_data:
            summary = self.comparison_data['summary']
            self.set_font('Arial', 'B', 11)
            self.cell(0, 8, 'Overall Statistics:', 0, 1)
            self.set_font('Arial', '', 10)
            
            stats = [
                ('Total Sheets', str(self.comparison_data['total_sheets'])),
                ('Total Columns Compared', str(summary['total_columns_compared'])),
                ('Matching Columns', str(summary['matching_columns'])),
                ('Different Columns', str(summary['different_columns'])),
                ('Columns with Errors', str(summary['error_columns'])),
                ('Success Rate', f"{summary['success_rate']}%")
            ]
            
            for label, value in stats:
                self.cell(60, 7, f"{label}:", 0, 0)
                self.set_font('Arial', 'B', 10)
                self.cell(30, 7, value, 0, 1)
                self.set_font('Arial', '', 10)
        
        self.ln(10)
    
    def sheets_section(self):
        self.chapter_title('DETAILED SHEET ANALYSIS')
        
        for sheet in self.comparison_data.get('sheets', []):
            # Sheet header
            self.set_font('Arial', 'B', 11)
            self.set_text_color(31, 73, 125)
            sheet_status = sheet.get('status', 'unknown').upper()
            status_color = {
                'processed': (0, 128, 0),  # Green
                'warning': (255, 165, 0),   # Orange
                'error': (255, 0, 0)        # Red
            }.get(sheet.get('status', '').lower(), (0, 0, 0))
            
            self.set_text_color(*status_color)
            self.cell(0, 8, f"Sheet: {sheet['sheet_name']} - Status: {sheet_status}", 0, 1)
            self.set_text_color(0, 0, 0)
            
            # Sheet summary
            self.set_font('Arial', '', 9)
            if sheet.get('error'):
                self.set_text_color(255, 0, 0)
                self.cell(0, 6, f"Error: {sheet['error']}", 0, 1)
                self.set_text_color(0, 0, 0)
            else:
                stats_text = f"Columns: Total({sheet['total_columns']}) | Matching({sheet['matching_columns']}) | Different({sheet['different_columns']}) | Errors({sheet.get('error_columns', 0)})"
                self.cell(0, 6, stats_text, 0, 1)
            
            self.ln(2)
            
            # Column details table
            if sheet['columns'] and len(sheet['columns']) > 0:
                self.column_details_table(sheet['columns'])
            
            self.ln(10)
    
    def column_details_table(self, columns):
        # Table header
        self.set_fill_color(200, 200, 200)
        self.set_font('Arial', 'B', 8)
        self.cell(60, 8, 'Column Name', 1, 0, 'C', 1)
        self.cell(25, 8, 'Type', 1, 0, 'C', 1)
        self.cell(25, 8, 'Status', 1, 0, 'C', 1)
        self.cell(80, 8, 'Details', 1, 1, 'C', 1)
        
        # Table rows
        self.set_font('Arial', '', 8)
        for col in columns:
            # Column name (truncate if too long)
            col_name = col['name'][:30] + '...' if len(col['name']) > 30 else col['name']
            self.cell(60, 6, col_name, 1, 0)
            
            # Type
            self.cell(25, 6, col.get('type', 'N/A'), 1, 0, 'C')
            
            # Status with color coding
            status = col.get('status', 'unknown')
            if status == 'matching':
                self.set_text_color(0, 128, 0)  # Green
            elif status == 'different':
                self.set_text_color(255, 165, 0)  # Orange
            elif status == 'error':
                self.set_text_color(255, 0, 0)  # Red
            else:
                self.set_text_color(0, 0, 0)  # Black
            
            self.cell(25, 6, status.upper(), 1, 0, 'C')
            self.set_text_color(0, 0, 0)  # Reset color
            
            # Details
            details = self.get_column_details(col)
            details = details[:50] + '...' if len(details) > 50 else details
            self.cell(80, 6, details, 1, 1)
    
    def get_column_details(self, col):
        status = col.get('status', 'unknown')
        
        if status == 'matching':
            if col.get('type') == 'numeric' and col.get('statistics'):
                stats = col['statistics']
                return f"Mean: {stats['mean']['file1']:.2f}"
            return "All values match"
        
        elif status == 'different':
            diffs = col.get('differences', [])
            if diffs:
                if col.get('type') == 'numeric':
                    return f"{len(diffs)} statistical differences"
                else:
                    return f"{len(diffs)} value count differences"
            return "Differences detected"
        
        elif status == 'error':
            return col.get('error', 'Unknown error')
        
        return "No details available"
    
    def differences_section(self):
        self.chapter_title('KEY DIFFERENCES HIGHLIGHT')
        
        significant_differences = []
        
        for sheet in self.comparison_data.get('sheets', []):
            if sheet.get('status') != 'processed':
                continue
                
            for col in sheet.get('columns', []):
                if col.get('status') == 'different' and col.get('differences'):
                    for diff in col['differences'][:3]:  # Show first 3 differences per column
                        significant_differences.append({
                            'sheet': sheet['sheet_name'],
                            'column': col['name'],
                            'difference': diff
                        })
        
        if not significant_differences:
            self.set_font('Arial', 'I', 10)
            self.cell(0, 8, "No significant differences found or all columns match perfectly.", 0, 1)
            return
        
        self.set_font('Arial', 'B', 9)
        self.cell(0, 8, "Top Differences Found:", 0, 1)
        self.ln(2)
        
        self.set_font('Arial', '', 8)
        for i, diff_info in enumerate(significant_differences[:10], 1):  # Limit to top 10
            sheet = diff_info['sheet']
            column = diff_info['column']
            difference = diff_info['difference']
            
            self.set_font('Arial', 'B', 8)
            self.cell(0, 6, f"{i}. {sheet} -> {column}", 0, 1)
            self.set_font('Arial', '', 8)
            
            if 'value' in difference:  # Text column difference
                self.cell(10, 5, '', 0, 0)
                self.cell(0, 5, f"Value '{difference['value']}': File1={difference['file1_count']}, File2={difference['file2_count']}", 0, 1)
            else:  # Numeric column difference
                self.cell(10, 5, '', 0, 0)
                self.cell(0, 5, f"{difference['statistic']}: File1={difference['file1_value']}, File2={difference['file2_value']}, Diff={difference['difference']}", 0, 1)
            
            self.ln(2)
    
    def generate_report(self):
        self.add_page()
        
        # Cover page
        self.set_font('Arial', 'B', 20)
        self.set_text_color(31, 73, 125)
        self.cell(0, 40, 'Excel File Comparison Report', 0, 1, 'C')
        self.ln(20)
        
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, f"{self.comparison_data['file1_name']}", 0, 1, 'C')
        self.set_font('Arial', '', 12)
        self.cell(0, 8, 'vs', 0, 1, 'C')
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, f"{self.comparison_data['file2_name']}", 0, 1, 'C')
        
        self.ln(30)
        self.set_font('Arial', 'I', 10)
        self.set_text_color(100, 100, 100)
        self.cell(0, 8, f"Generated on {self.comparison_data['comparison_time']}", 0, 1, 'C')
        
        # Add content pages
        self.add_page()
        self.summary_section()
        self.sheets_section()
        self.differences_section()

def generate_pdf_report(comparison_data, output_path):
    """Generate PDF report from comparison data"""
    try:
        pdf = PDFReport(comparison_data)
        pdf.generate_report()
        pdf.output(output_path)
        return True
    except Exception as e:
        print(f"Error generating PDF report: {str(e)}")
        return False
