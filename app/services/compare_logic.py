import pandas as pd
import numpy as np
import os
from datetime import datetime
import traceback
import traceback
import logging
from io import BytesIO
import time


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_FOLDER = os.path.join(BASE_DIR, "reports")
os.makedirs(REPORT_FOLDER, exist_ok=True)

ALLOWED_EXT = {".xls", ".xlsx"}

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(BASE_DIR, 'comparison.log')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_FOLDER = os.path.join(BASE_DIR, "reports")
os.makedirs(REPORT_FOLDER, exist_ok=True)

ALLOWED_EXT = {".xls", ".xlsx"}

def allowed_filename(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT

def safe_parse_excel_from_memory(file_stream, sheet_name):
    """Parse Excel sheet directly from memory without saving to disk"""
    try:
        start_time = time.time()
        df = pd.read_excel(file_stream, sheet_name=sheet_name, engine='openpyxl')
        logger.info(f"Parsed sheet '{sheet_name}' in {time.time() - start_time:.2f}s, shape: {df.shape}")
        
        if df.empty or len(df) == 0:
            return pd.DataFrame()
        return df
    except Exception as e:
        logger.warning(f"Could not parse sheet '{sheet_name}': {str(e)}")
        return pd.DataFrame()

def efficient_column_comparison(col1, col2, col_name, force_object_cols):
    """Optimized column comparison with performance improvements"""
    start_time = time.time()
    
    try:
        col_data = {
            "name": col_name,
            "type": "unknown",
            "status": "error",
            "differences": [],
            "error": None
        }
        
        # Quick empty checks
        if col1.empty and col2.empty:
            col_data.update({"type": "empty", "status": "matching", "error": "Both columns empty"})
            return col_data
            
        if col1.empty or col2.empty:
            col_data.update({"type": "empty", "status": "different", 
                           "error": f"One column empty (col1: {len(col1)}, col2: {len(col2)})"})
            return col_data

        # Fast null checks using numpy
        col1_null_mask = col1.isna()
        col2_null_mask = col2.isna()
        
        if col1_null_mask.all() and col2_null_mask.all():
            col_data.update({"type": "null", "status": "matching", "error": "Both columns null"})
            return col_data
            
        if col1_null_mask.all() or col2_null_mask.all():
            col_data.update({"type": "null", "status": "different", "error": "One column null"})
            return col_data

        # Determine column type - optimized check
        is_numeric = False
        try:
            # Sample first few rows for type detection instead of checking entire column
            sample_size = min(100, len(col1), len(col2))
            col1_sample = col1.head(sample_size).dropna()
            col2_sample = col2.head(sample_size).dropna()
            
            if len(col1_sample) > 0 and len(col2_sample) > 0:
                is_numeric = (pd.api.types.is_numeric_dtype(col1_sample) and 
                            pd.api.types.is_numeric_dtype(col2_sample) and 
                            col_name not in force_object_cols)
        except Exception:
            is_numeric = False

        col_data["type"] = "numeric" if is_numeric else "text"

        if col_name in force_object_cols or not is_numeric:
            # Optimized text comparison
            try:
                # Use value_counts with dropna=False for better performance
                vc1 = col1.astype(str).value_counts(dropna=False)
                vc2 = col2.astype(str).value_counts(dropna=False)
                
                # Find differences efficiently using set operations
                all_values = set(vc1.index).union(set(vc2.index))
                diffs = []
                
                for val in all_values:
                    count1 = vc1.get(val, 0)
                    count2 = vc2.get(val, 0)
                    if count1 != count2:
                        diffs.append({
                            "value": str(val),
                            "file1_count": int(count1),
                            "file2_count": int(count2)
                        })
                    if len(diffs) >= 10:  # Early stop if we have enough differences
                        break

                col_data.update({
                    "status": "different" if diffs else "matching",
                    "differences": diffs
                })
                
            except Exception as e:
                logger.warning(f"Text comparison failed for {col_name}: {str(e)}")
                col_data.update({"status": "error", "error": f"Text comparison failed: {str(e)}"})
                
        else:
            # Optimized numeric comparison
            try:
                # Convert to numeric in one go
                col1_numeric = pd.to_numeric(col1, errors='coerce')
                col2_numeric = pd.to_numeric(col2, errors='coerce')
                
                # Quick check for failed conversion
                if col1_numeric.isna().all() or col2_numeric.isna().all():
                    col_data.update({"type": "text", "status": "error", 
                                   "error": "Failed numeric conversion"})
                    return col_data

                # Calculate only essential statistics
                stats = {}
                differences_found = []
                
                # Use numpy for faster calculations
                col1_vals = col1_numeric.values
                col2_vals = col2_numeric.values
                
                # Remove NaN values for statistical calculations
                col1_clean = col1_vals[~np.isnan(col1_vals)]
                col2_clean = col2_vals[~np.isnan(col2_vals)]
                
                if len(col1_clean) > 0 and len(col2_clean) > 0:
                    stats = {
                        "sum": {"file1": float(np.sum(col1_clean)), "file2": float(np.sum(col2_clean))},
                        "mean": {"file1": float(np.mean(col1_clean)), "file2": float(np.mean(col2_clean))},
                        "min": {"file1": float(np.min(col1_clean)), "file2": float(np.min(col2_clean))},
                        "max": {"file1": float(np.max(col1_clean)), "file2": float(np.max(col2_clean))},
                    }
                    
                    # Check differences with vectorized operations
                    for stat, values in stats.items():
                        v1, v2 = values["file1"], values["file2"]
                        
                        if np.isnan(v1) or np.isnan(v2) or np.isinf(v1) or np.isinf(v2):
                            differences_found.append({
                                "statistic": stat,
                                "file1_value": v1,
                                "file2_value": v2,
                                "difference": "NaN/Inf detected"
                            })
                        elif stat in ["mean", "sum"]:
                            if abs(v1) > 1 or abs(v2) > 1:
                                relative_diff = abs(v1 - v2) / max(abs(v1), abs(v2))
                                if relative_diff > 1e-6:  # Slightly relaxed tolerance
                                    differences_found.append({
                                        "statistic": stat,
                                        "file1_value": round(v1, 4),
                                        "file2_value": round(v2, 4),
                                        "difference": round(v2 - v1, 4)
                                    })
                            else:
                                if abs(v1 - v2) > 1e-6:
                                    differences_found.append({
                                        "statistic": stat,
                                        "file1_value": round(v1, 4),
                                        "file2_value": round(v2, 4),
                                        "difference": round(v2 - v1, 4)
                                    })
                        elif stat in ["min", "max"]:
                            if abs(v1 - v2) > 1e-6:
                                differences_found.append({
                                    "statistic": stat,
                                    "file1_value": v1,
                                    "file2_value": v2,
                                    "difference": round(v2 - v1, 4)
                                })

                col_data.update({
                    "status": "different" if differences_found else "matching",
                    "differences": differences_found,
                    "statistics": stats
                })
                    
            except Exception as e:
                logger.warning(f"Numeric comparison failed for {col_name}: {str(e)}")
                col_data.update({"status": "error", "error": f"Numeric comparison failed: {str(e)}"})

        logger.debug(f"Column {col_name} comparison completed in {time.time() - start_time:.3f}s")
        return col_data

    except Exception as e:
        logger.error(f"Unexpected error in column comparison {col_name}: {str(e)}")
        return {
            "name": col_name,
            "type": "unknown",
            "status": "error",
            "differences": [],
            "error": f"Unexpected error: {str(e)}"
        }

def compare_excel_stats(file1, file2):
    """Optimized Excel comparison without temporary file operations"""
    start_time = time.time()
    logger.info(f"Starting comparison: {file1.filename} vs {file2.filename}")
    
    try:
        # Read Excel files directly from memory
        file1_stream = BytesIO(file1.read())
        file2_stream = BytesIO(file2.read())
        
        # Reset stream positions for multiple reads
        file1_stream.seek(0)
        file2_stream.seek(0)
        
        # Get sheet names
        try:
            xl1 = pd.ExcelFile(file1_stream, engine='openpyxl')
            xl2 = pd.ExcelFile(file2_stream, engine='openpyxl')
            
            sheets1 = xl1.sheet_names
            sheets2 = xl2.sheet_names
            common_sheets = sorted(set(sheets1).intersection(sheets2))
            
            logger.info(f"Found {len(common_sheets)} common sheets: {common_sheets}")
            
        except Exception as e:
            logger.error(f"Failed to read Excel files: {str(e)}")
            return {"error": f"Failed to read Excel files: {str(e)}"}

        force_object_cols = {"UW_Year", "Loss_Period"}
        
        comparison_results = {
            "file1_name": file1.filename,
            "file2_name": file2.filename,
            "comparison_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_sheets": len(common_sheets),
            "sheets_processed": 0,
            "sheets_failed": 0,
            "sheets": []
        }

        if not common_sheets:
            logger.warning("No common sheets found between files")
            comparison_results["warning"] = "No common sheets found between files"
            return comparison_results

        # Process sheets
        for sheet_idx, sheet in enumerate(common_sheets):
            sheet_start_time = time.time()
            logger.info(f"Processing sheet {sheet_idx + 1}/{len(common_sheets)}: {sheet}")
            
            sheet_data = {
                "sheet_name": sheet,
                "status": "processed",
                "error": None,
                "total_columns": 0,
                "matching_columns": 0,
                "different_columns": 0,
                "error_columns": 0,
                "columns": []
            }

            try:
                # Reset streams for each sheet read
                file1_stream.seek(0)
                file2_stream.seek(0)
                
                # Parse sheets
                df1 = safe_parse_excel_from_memory(file1_stream, sheet)
                df2 = safe_parse_excel_from_memory(file2_stream, sheet)
                
                if df1.empty or df2.empty:
                    sheet_data.update({
                        "status": "warning",
                        "error": "One or both sheets are empty",
                        "total_columns": 0
                    })
                    comparison_results["sheets"].append(sheet_data)
                    continue

                # Get common columns
                common_cols = list(df1.columns.intersection(df2.columns))
                sheet_data["total_columns"] = len(common_cols)
                
                if not common_cols:
                    sheet_data.update({
                        "status": "warning",
                        "error": "No common columns found"
                    })
                    comparison_results["sheets"].append(sheet_data)
                    continue

                logger.info(f"Sheet {sheet}: Comparing {len(common_cols)} columns")
                
                # Process columns in batches for better memory management
                for col_idx, col in enumerate(common_cols):
                    try:
                        col_result = efficient_column_comparison(
                            df1[col], df2[col], col, force_object_cols
                        )
                        
                        # Update counters
                        status = col_result.get("status")
                        if status == "matching":
                            sheet_data["matching_columns"] += 1
                        elif status == "different":
                            sheet_data["different_columns"] += 1
                        elif status == "error":
                            sheet_data["error_columns"] += 1
                            
                        sheet_data["columns"].append(col_result)
                        
                    except Exception as col_error:
                        logger.warning(f"Column {col} failed: {str(col_error)}")
                        sheet_data["columns"].append({
                            "name": col, "type": "unknown", "status": "error",
                            "differences": [], "error": f"Processing failed: {str(col_error)}"
                        })
                        sheet_data["error_columns"] += 1

                comparison_results["sheets_processed"] += 1
                logger.info(f"Sheet {sheet} completed in {time.time() - sheet_start_time:.2f}s")
                
            except Exception as sheet_error:
                logger.error(f"Sheet {sheet} failed: {str(sheet_error)}")
                sheet_data.update({
                    "status": "error",
                    "error": f"Sheet processing failed: {str(sheet_error)}",
                    "total_columns": 0,
                    "columns": []
                })
                comparison_results["sheets_failed"] += 1

            comparison_results["sheets"].append(sheet_data)

        # Calculate summary
        total_columns = sum(sheet["total_columns"] for sheet in comparison_results["sheets"])
        matching_columns = sum(sheet["matching_columns"] for sheet in comparison_results["sheets"])
        different_columns = sum(sheet["different_columns"] for sheet in comparison_results["sheets"])
        error_columns = sum(sheet["error_columns"] for sheet in comparison_results["sheets"])
        
        comparison_results["summary"] = {
            "total_columns_compared": total_columns,
            "matching_columns": matching_columns,
            "different_columns": different_columns,
            "error_columns": error_columns,
            "success_rate": round((matching_columns + different_columns) / total_columns * 100, 2) if total_columns > 0 else 0
        }

        total_time = time.time() - start_time
        logger.info(f"Comparison completed in {total_time:.2f}s. "
                   f"Results: {matching_columns} matching, {different_columns} different, {error_columns} error columns")
        
        return comparison_results

    except Exception as e:
        error_msg = f"Critical error in comparison: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return {"error": error_msg}