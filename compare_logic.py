import pandas as pd

def compare_excel_stats(file1, file2):
    print("Comparing files:", file1, file2)
    xl1 = pd.ExcelFile(file1)
    xl2 = pd.ExcelFile(file2)

    common_sheets = sorted(set(xl1.sheet_names).intersection(xl2.sheet_names))
    print("Common sheets:", common_sheets)

    force_object_cols = {"UW_Year", "Loss_Period"}

    result = []
    if not common_sheets:
        return [{"sheet": None, "message": "❌ No common sheets found."}]

    for sheet in common_sheets:
        df1 = xl1.parse(sheet)
        df2 = xl2.parse(sheet)
        common_cols = list(df1.columns.intersection(df2.columns))
        print(f"Sheet '{sheet}' - common columns:", common_cols)
        if not common_cols:
            result.append({"sheet": sheet, "message": "⚠️ No common columns found."})
            continue

        df1 = df1[common_cols]
        df2 = df2[common_cols]
        sheet_report = [f"📄 Sheet: {sheet}"]

        for col in common_cols:
            col1 = df1[col]
            col2 = df2[col]
            if col in force_object_cols or not (
                pd.api.types.is_numeric_dtype(col1)
                and pd.api.types.is_numeric_dtype(col2)
            ):
                vc1 = col1.astype(str).value_counts().sort_index()
                vc2 = col2.astype(str).value_counts().sort_index()
                combined_index = vc1.index.union(vc2.index)

                diffs = [
                    (val, vc1.get(val, 0), vc2.get(val, 0))
                    for val in combined_index
                    if vc1.get(val, 0) != vc2.get(val, 0)
                ]

                if diffs:
                    sheet_report.append(f"⚠️ Column: {col} - Value count differences:")
                    for val, c1, c2 in diffs[:10]:
                        sheet_report.append(f"   '{val}': File1={c1}, File2={c2}")
                else:
                    sheet_report.append(f"✅ Column: {col} matches.")
            else:
                col1 = pd.to_numeric(col1, errors="coerce")
                col2 = pd.to_numeric(col2, errors="coerce")

                summary = {
                    "sum": (col1.sum(), col2.sum()),
                    "mean": (col1.mean(), col2.mean()),
                    "min": (col1.min(), col2.min()),
                    "max": (col1.max(), col2.max()),
                }    
            
                differences = {
                    # stat: (round(v1, 4), round(v2, 4)) 
                    # for stat, (v1, v2) in summary.items()
                    # if stat in (["mean", "sum"] and round(v1, 4) != round(v2, 4))
                }

                for stat, (v1, v2) in summary.items():
                    if stat in ["mean", "sum"] and round(v1, 4) != round(v2, 4):
                        differences[stat] = (round(v1, 4), round(v2, 4))
                    elif stat in ["min", "max"] and v1 != v2:
                        differences[stat] = (v1, v2) if type(v1) != float or type(v2) != float else (round(v1, 4), round(v2, 4))



                if differences:
                    for stat, (v1, v2) in differences.items():
                        sheet_report.append(f"⚠️ {col}: {stat} differs (F1={v1}, F2={v2})")
                else:
                    sheet_report.append(f"✅ {col} has no statistical differences.")

        result.append({"sheet": sheet, "message": "\n".join(sheet_report)})

    return result
