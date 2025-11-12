import numpy as np
def format_number(value, precision=2, multiply_factor=1, percentage_sign=False, add_commas=True):
    """
    Format numerical values for better viewing with comma separation and smart precision.
    
    Args:
        value: The numerical value to format
        precision: Number of decimal places (default: 2)
        multiply_factor: Factor to multiply the value by (default: 1)
        percentage_sign: Whether to add percentage sign (default: False)
        add_commas: Whether to add thousand separators (default: True)
    
    Returns:
        Formatted string representation of the number
    """
    # Handle None, strings, and zero values
    if value is None or isinstance(value, str):
        return str(value) if value is not None else "N/A"
    
    # Handle NaN and infinite values
    if isinstance(value, (int, float)) and (np.isnan(value) or np.isinf(value)):
        if np.isnan(value):
            return "NaN"
        elif value > 0:
            return "∞"
        else:
            return "-∞"
    
    # Apply multiplication factor
    try:
        multiplied_value = float(value) * multiply_factor
    except (TypeError, ValueError):
        return str(value)
    
    # Handle very small numbers (avoid scientific notation)
    if abs(multiplied_value) > 0 and abs(multiplied_value) < 0.0001:
        return f"{multiplied_value:.2e}"
    
    # Determine optimal precision for the value
    actual_precision = precision
    if abs(multiplied_value) >= 1000:
        # For large numbers, reduce precision
        actual_precision = max(0, precision - 1)
    elif abs(multiplied_value) < 1 and multiplied_value != 0:
        # For small numbers, increase precision
        actual_precision = precision + 2
    
    # Format the number
    if percentage_sign:
        formatted_value = f"{multiplied_value:.{actual_precision}f}%"
    else:
        formatted_value = f"{multiplied_value:.{actual_precision}f}"
    
    # Add comma separators for thousands
    if add_commas and not ('e' in formatted_value.lower() or '∞' in formatted_value or 'nan' in formatted_value.lower()):
        try:
            # Split into integer and decimal parts
            if '.' in formatted_value:
                int_part, decimal_part = formatted_value.split('.')
            else:
                int_part, decimal_part = formatted_value, ''
            
            # Add commas to integer part
            if int_part.startswith('-'):
                sign = '-'
                int_part = int_part[1:]
            else:
                sign = ''
            
            # Add thousand separators
            int_with_commas = ''
            for i, digit in enumerate(reversed(int_part)):
                if i > 0 and i % 3 == 0:
                    int_with_commas = ',' + int_with_commas
                int_with_commas = digit + int_with_commas
            
            # Reconstruct the number
            if decimal_part:
                formatted_value = f"{sign}{int_with_commas}.{decimal_part}"
            else:
                formatted_value = f"{sign}{int_with_commas}"
                
        except Exception:
            # If comma formatting fails, return the original formatted value
            pass
    
    return formatted_value

def format_dictionary(d, precision=2, multiply_factor=1, percentage_sign=False, add_commas=True):
    """
    Format all numerical values in a dictionary.
    
    Args:
        d: Dictionary containing numerical values
        precision: Number of decimal places
        multiply_factor: Factor to multiply values by
        percentage_sign: Whether to add percentage sign
        add_commas: Whether to add thousand separators
    
    Returns:
        New dictionary with formatted values
    """
    if not isinstance(d, dict):
        return d
    
    formatted_dict = {}
    for k, v in d.items():
        if isinstance(v, (int, float)):
            formatted_dict[k] = format_number(v, precision, multiply_factor, percentage_sign, add_commas)
        elif isinstance(v, dict):
            formatted_dict[k] = format_dictionary(v, precision, multiply_factor, percentage_sign, add_commas)
        elif isinstance(v, list):
            formatted_dict[k] = [format_number(item, precision, multiply_factor, percentage_sign, add_commas) 
                               if isinstance(item, (int, float)) else item for item in v]
        else:
            formatted_dict[k] = v
    return formatted_dict

def format_statistics_display(statistics_dict, precision=2):
    """
    Specialized formatter for statistics display in the comparison results.
    
    Args:
        statistics_dict: Dictionary containing statistical values
        precision: Number of decimal places
    
    Returns:
        Formatted statistics dictionary
    """
    if not isinstance(statistics_dict, dict):
        return statistics_dict
    
    formatted_stats = {}
    for stat_name, files_dict in statistics_dict.items():
        if isinstance(files_dict, dict) and 'file1' in files_dict and 'file2' in files_dict:
            # Adjust precision based on the statistic type
            stat_precision = precision
            if stat_name in ['sum', 'mean']:
                stat_precision = max(2, precision)
            elif stat_name in ['min', 'max']:
                stat_precision = precision
            
            formatted_stats[stat_name] = {
                'file1': format_number(files_dict['file1'], stat_precision),
                'file2': format_number(files_dict['file2'], stat_precision)
            }
        else:
            formatted_stats[stat_name] = files_dict
    
    return formatted_stats

def format_comparison_results(comparison_results, precision=2):
    """
    Format the entire comparison results for better display.
    
    Args:
        comparison_results: The complete comparison results dictionary
        precision: Number of decimal places for formatting
    
    Returns:
        Formatted comparison results
    """
    if not isinstance(comparison_results, dict):
        return comparison_results
    
    formatted_results = comparison_results.copy()
    
    # Format summary statistics
    if 'summary' in formatted_results:
        formatted_results['summary'] = format_dictionary(
            formatted_results['summary'], precision=0, add_commas=True
        )
    
    # Format sheet statistics and column data
    if 'sheets' in formatted_results:
        for sheet in formatted_results['sheets']:
            if 'columns' in sheet:
                for column in sheet['columns']:
                    # Format statistics for numeric columns
                    if 'statistics' in column and column.get('type') == 'numeric':
                        column['statistics'] = format_statistics_display(
                            column['statistics'], precision=precision
                        )
                    
                    # Format differences
                    if 'differences' in column and column['differences']:
                        formatted_differences = []
                        for diff in column['differences']:
                            formatted_diff = {}
                            for key, value in diff.items():
                                if isinstance(value, (int, float)):
                                    # Use different precision for different types of values
                                    if key in ['file1_value', 'file2_value', 'difference']:
                                        diff_precision = precision
                                        if abs(value) < 0.01 and value != 0:
                                            diff_precision = 6  # More precision for very small differences
                                        formatted_diff[key] = format_number(value, diff_precision)
                                    else:
                                        formatted_diff[key] = value
                                else:
                                    formatted_diff[key] = value
                            formatted_differences.append(formatted_diff)
                        column['differences'] = formatted_differences
    
    return formatted_results

# Example usage and test cases
if __name__ == "__main__":
    # Test the formatting functions
    test_cases = [
        1234567.89123,           # Large number with decimals
        0.00012345,              # Very small number
        1234.0,                  # Integer-like float
        0,                       # Zero
        999999.999,              # Borderline large number
        0.123456789,             # Medium precision
        -123456.78,              # Negative number
        float('nan'),            # NaN
        float('inf'),            # Infinity
        None,                    # None value
        "string_value",          # String
    ]
    
    print("Number Formatting Examples:")
    print("-" * 50)
    for test_value in test_cases:
        formatted = format_number(test_value, precision=2)
        print(f"{test_value!r:>20} -> {formatted}")
    
    # Test dictionary formatting
    test_dict = {
        "revenue": 1234567.89,
        "growth_rate": 0.15432,
        "count": 10000,
        "small_value": 0.000456,
        "nested": {
            "profit": 456789.12,
            "margin": 0.2345
        }
    }
    
    print("\nDictionary Formatting Examples:")
    print("-" * 50)
    formatted_dict = format_dictionary(test_dict, precision=2)
    print(formatted_dict)
    
    # Test statistics formatting
    stats_test = {
        "sum": {"file1": 1234567.891, "file2": 1234567.892},
        "mean": {"file1": 1234.56789, "file2": 1234.56788},
        "min": {"file1": 0.000123, "file2": 0.000124},
        "max": {"file1": 999999.999, "file2": 999999.998}
    }
    
    print("\nStatistics Formatting Examples:")
    print("-" * 50)
    formatted_stats = format_statistics_display(stats_test, precision=3)
    print(formatted_stats)