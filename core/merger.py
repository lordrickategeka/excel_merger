import pandas as pd
from collections import defaultdict

class ExcelColumnMerger:
    """
    Core class for Excel column merging and analysis operations.
    """
    def __init__(self):
        self.input_file = None
        self.output_file = None
        self.current_sheets = {}  # Store current dataframes for each sheet
        self.modified_sheets = set()
        
    def set_input_file(self, file_path):
        """Set the input file and read its contents"""
        self.input_file = file_path
        self.read_excel_file()
        
    def read_excel_file(self):
        """Read all sheets from the Excel file"""
        if not self.input_file:
            return False
            
        try:
            # Read all sheets
            xl = pd.ExcelFile(self.input_file)
            self.current_sheets = {
                sheet: pd.read_excel(self.input_file, sheet_name=sheet)
                for sheet in xl.sheet_names
            }
            return True
        except Exception as e:
            raise Exception(f"Failed to read file: {str(e)}")
    
    def analyze_file(self):
        """Analyze the Excel file for duplicate column names (case-insensitive)"""
        if not self.input_file or not self.current_sheets:
            return None
            
        try:
            sheets_data = {}
            
            for sheet_name, df in self.current_sheets.items():
                # Check for duplicate columns (case-insensitive)
                column_groups = defaultdict(list)
                
                # Group columns by lowercase name
                for col in df.columns:
                    column_groups[str(col).lower()].append(col)
                
                # Keep only groups with more than one column
                duplicate_columns = {k: v for k, v in column_groups.items() if len(v) > 1}
                
                if duplicate_columns:
                    sheets_data[sheet_name] = {
                        'dataframe': df,
                        'duplicate_columns': duplicate_columns
                    }
            
            if not sheets_data:
                return {"status": "no_duplicates"}
            
            return {
                "status": "ok",
                "sheets_data": sheets_data,
                "sheet_names": list(sheets_data.keys())
            }
            
        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }
    
    def merge_columns(self, analysis, strategy="first_non_empty"):
        """Merge duplicate columns based on the selected strategy"""
        if not analysis or analysis["status"] != "ok":
            return False
            
        try:
            sheets_data = analysis["sheets_data"]
            
            for sheet_name, sheet_data in sheets_data.items():
                df = sheet_data['dataframe']
                duplicate_columns = sheet_data['duplicate_columns']
                
                # Process each group of duplicate columns
                for base_name, columns in duplicate_columns.items():
                    # Choose merge strategy
                    if strategy == "first_non_empty":
                        # Create a new column combining non-empty values
                        new_values = df[columns].apply(
                            lambda row: next((x for x in row if pd.notna(x)), None), 
                            axis=1
                        )
                    elif strategy == "sum":
                        # Sum numeric values, ignoring non-numeric
                        numeric_columns = []
                        for col in columns:
                            if pd.api.types.is_numeric_dtype(df[col]):
                                numeric_columns.append(col)
                        
                        if numeric_columns:
                            new_values = df[numeric_columns].sum(axis=1)
                        else:
                            # If no numeric columns, use first_non_empty
                            new_values = df[columns].apply(
                                lambda row: next((x for x in row if pd.notna(x)), None), 
                                axis=1
                            )
                    elif strategy == "concatenate":
                        # Concatenate all non-empty string values
                        new_values = df[columns].apply(
                            lambda row: " ".join([str(x) for x in row if pd.notna(x) and str(x).strip() != ""]), 
                            axis=1
                        )
                    
                    # Create a new dataframe without the duplicate columns
                    new_df = df.drop(columns=columns)
                    
                    # Add the merged column (use the first duplicate name)
                    new_df[columns[0]] = new_values
                    
                    # Update the dataframe
                    df = new_df
                
                # Update the current sheet
                self.current_sheets[sheet_name] = df
            
            return True
            
        except Exception as e:
            raise Exception(f"Failed to merge columns: {str(e)}")
    
    def manual_merge_columns(self, sheet_name, columns, new_column_name, strategy="first_non_empty", delete_source=True):
        """
        Merge selected columns in a sheet based on a specified strategy.
        
        Args:
            sheet_name: Name of the sheet to modify
            columns: List of column names to merge
            new_column_name: Name for the merged column
            strategy: Merge strategy ('first_non_empty', 'sum', 'concatenate', or 'stack_values')
            delete_source: Whether to delete source columns after merging
            
        Returns:
            bool: Success or failure
        """
        try:
            if sheet_name not in self.current_sheets:
                return False
            
            # Special handling for stack_values strategy
            if strategy == "stack_values":
                return self.stack_values_merge(sheet_name, columns, new_column_name, delete_source)
            
            # Get the dataframe
            df = self.current_sheets[sheet_name]
            
            # Check if all selected columns exist
            if not all(col in df.columns for col in columns):
                return False
            
            # Apply the appropriate merge strategy
            if strategy == "first_non_empty":
                # Combine columns keeping the first non-empty value for each row
                df[new_column_name] = df[columns].apply(
                    lambda row: next((v for v in row if pd.notna(v) and str(v).strip() != ""), None), 
                    axis=1
                )
            elif strategy == "sum":
                # Helper function to convert to numeric safely
                def safe_numeric(val):
                    try:
                        return float(val) if pd.notna(val) and str(val).strip() != "" else 0
                    except (ValueError, TypeError):
                        return 0
                
                # Sum numeric values, handling non-numeric values as 0
                df[new_column_name] = df[columns].apply(
                    lambda row: sum(safe_numeric(v) for v in row), 
                    axis=1
                )
            elif strategy == "concatenate":
                # Concatenate non-empty values with separator
                separator = " | "  # Can be customized if needed
                df[new_column_name] = df[columns].apply(
                    lambda row: separator.join(str(v) for v in row if pd.notna(v) and str(v).strip() != ""), 
                    axis=1
                )
            else:
                # Unknown strategy
                return False
            
            # Delete source columns if requested
            if delete_source:
                df = df.drop(columns=columns)
            
            # Update the dataframe in our dictionary
            self.current_sheets[sheet_name] = df
            
            # Mark that the sheet was modified
            self.modified_sheets.add(sheet_name)
            
            return True
            
        except Exception as e:
            print(f"Error in manual_merge_columns: {str(e)}")
            return False
    
    def compare_columns_for_duplicates(self, sheet_name, columns_to_compare):
        """Compare values across selected columns and identify duplicate values"""
        if not sheet_name or not columns_to_compare or sheet_name not in self.current_sheets:
            return None
        
        try:
            df = self.current_sheets[sheet_name]
            
            # Ensure all columns exist
            for col in columns_to_compare:
                if col not in df.columns:
                    raise Exception(f"Column '{col}' not found in sheet '{sheet_name}'")
            
            # Create a result dataframe
            result_df = pd.DataFrame(index=df.index)
            
            # Add the columns we're comparing
            for col in columns_to_compare:
                result_df[col] = df[col]
            
            # Find rows with duplicate values across columns
            duplicates_mask = pd.DataFrame(index=df.index)
            
            # Compare each pair of columns
            for i, col1 in enumerate(columns_to_compare):
                for col2 in columns_to_compare[i+1:]:
                    # Check for exact matches (excluding NaN)
                    mask = (df[col1] == df[col2]) & df[col1].notna()
                    duplicates_mask[f"{col1}_vs_{col2}"] = mask
            
            # Combine all masks to find any row with at least one duplicate
            has_duplicates = duplicates_mask.any(axis=1)
            result_df['Has_Duplicates'] = has_duplicates
            
            # Add information about which columns have duplicates
            for i, col1 in enumerate(columns_to_compare):
                for col2 in columns_to_compare[i+1:]:
                    dup_col_name = f"{col1}_eq_{col2}"
                    result_df[dup_col_name] = (df[col1] == df[col2]) & df[col1].notna()
            
            return {
                'result_df': result_df,
                'duplicate_rows': result_df[has_duplicates],
                'duplicate_count': has_duplicates.sum(),
                'total_rows': len(df)
            }
            
        except Exception as e:
            raise Exception(f"Failed to compare columns: {str(e)}")
    
    def create_common_column(self, sheet_name, columns_to_combine, new_column_name, strategy="first_non_empty", mark_duplicates=False):
        """Create a common column from selected columns and optionally mark duplicate values"""
        if not sheet_name or not columns_to_combine or sheet_name not in self.current_sheets:
            return False
        
        try:
            df = self.current_sheets[sheet_name]
            
            # Ensure all columns exist
            for col in columns_to_combine:
                if col not in df.columns:
                    raise Exception(f"Column '{col}' not found in sheet '{sheet_name}'")
            
            # First identify duplicates
            comparison_result = self.compare_columns_for_duplicates(sheet_name, columns_to_combine)
            if not comparison_result:
                return False
            
            # Apply merge strategy
            if strategy == "first_non_empty":
                new_values = df[columns_to_combine].apply(
                    lambda row: next((x for x in row if pd.notna(x)), None), 
                    axis=1
                )
            elif strategy == "prioritize_duplicates":
                # For each row, prioritize values that appear in multiple columns
                def get_duplicate_value(row):
                    values = [row[col] for col in columns_to_combine if pd.notna(row[col])]
                    if not values:
                        return None
                    
                    # Count occurrences of each value
                    value_counts = {}
                    for val in values:
                        value_counts[val] = value_counts.get(val, 0) + 1
                    
                    # Get the value with the highest count
                    max_count = max(value_counts.values())
                    if max_count > 1:  # If there are duplicates
                        for val, count in value_counts.items():
                            if count == max_count:
                                return val
                    
                    # If no duplicates, use first non-empty
                    return values[0]
                
                new_values = pd.Series([get_duplicate_value(row) for _, row in df[columns_to_combine].iterrows()], index=df.index)
                
            elif strategy == "mark_duplicates":
                # Similar to first_non_empty but mark duplicate values
                def mark_if_duplicate(row):
                    values = [row[col] for col in columns_to_combine if pd.notna(row[col])]
                    if not values:
                        return None
                    
                    # Count occurrences of each value
                    value_counts = {}
                    for val in values:
                        value_counts[val] = value_counts.get(val, 0) + 1
                    
                    # Get the first non-empty value
                    value = values[0]
                    
                    # If it appears multiple times, mark it
                    if value_counts[value] > 1:
                        return f"{value} (duplicate)"
                    return value
                
                new_values = pd.Series([mark_if_duplicate(row) for _, row in df[columns_to_combine].iterrows()], index=df.index)
                
            elif strategy == "concatenate":
                new_values = df[columns_to_combine].apply(
                    lambda row: " | ".join([str(x) for x in row if pd.notna(x) and str(x).strip() != ""]), 
                    axis=1
                )
            
            # Add the new column
            df[new_column_name] = new_values
            
            # If requested, add a column that marks which rows have duplicates
            if mark_duplicates:
                df[f"{new_column_name}_has_duplicate"] = comparison_result['result_df']['Has_Duplicates']
            
            # Update the current sheet
            self.current_sheets[sheet_name] = df
            
            return True
            
        except Exception as e:
            raise Exception(f"Failed to create common column: {str(e)}")
    
    def analyze_column(self, sheet_name, column_name):
        """Analyze a single column for data statistics"""
        if not sheet_name or not column_name or sheet_name not in self.current_sheets:
            return None
        
        try:
            df = self.current_sheets[sheet_name]
            
            # Ensure column exists
            if column_name not in df.columns:
                raise Exception(f"Column '{column_name}' not found in sheet '{sheet_name}'")
            
            # Get the column data
            column_data = df[column_name]
            
            # Basic statistics
            total_rows = len(column_data)
            non_empty_count = column_data.notna().sum()
            empty_count = total_rows - non_empty_count
            
            # Get unique values
            unique_values = column_data.dropna().unique()
            unique_count = len(unique_values)
            
            # Sample values (first 10)
            sample_values = column_data.dropna().head(10).tolist()
            
            # Determine data type
            data_type = "Mixed"
            if pd.api.types.is_numeric_dtype(column_data):
                data_type = "Numeric"
                # Add numeric stats
                numeric_stats = {
                    "min": column_data.min() if non_empty_count > 0 else None,
                    "max": column_data.max() if non_empty_count > 0 else None,
                    "mean": column_data.mean() if non_empty_count > 0 else None,
                    "median": column_data.median() if non_empty_count > 0 else None
                }
            elif pd.api.types.is_string_dtype(column_data):
                data_type = "Text"
                # Add text stats
                text_lengths = column_data.dropna().astype(str).str.len()
                text_stats = {
                    "min_length": text_lengths.min() if non_empty_count > 0 else None,
                    "max_length": text_lengths.max() if non_empty_count > 0 else None,
                    "avg_length": text_lengths.mean() if non_empty_count > 0 else None
                }
                numeric_stats = None
            elif pd.api.types.is_datetime64_dtype(column_data):
                data_type = "Date/Time"
                # Add date stats
                numeric_stats = {
                    "earliest": column_data.min() if non_empty_count > 0 else None,
                    "latest": column_data.max() if non_empty_count > 0 else None
                }
                text_stats = None
            else:
                # For mixed or other types
                text_stats = None
                numeric_stats = None
            
            # Return the analysis result
            result = {
                "total_rows": total_rows,
                "non_empty_count": non_empty_count,
                "empty_count": empty_count,
                "empty_percentage": (empty_count / total_rows * 100) if total_rows > 0 else 0,
                "unique_count": unique_count,
                "data_type": data_type,
                "sample_values": sample_values
            }
            
            # Add type-specific stats if available
            if data_type == "Numeric":
                result["numeric_stats"] = numeric_stats
            elif data_type == "Text":
                result["text_stats"] = text_stats
            elif data_type == "Date/Time":
                result["date_stats"] = numeric_stats
            
            return result
            
        except Exception as e:
            raise Exception(f"Failed to analyze column: {str(e)}")
        
    def get_non_empty_columns(self, sheet_name):
        """
        Get a list of columns that contain data (not completely empty)
        
        Args:
            sheet_name: Name of the sheet to check
            
        Returns:
            List of column names that contain at least one non-empty value
        """
        if not sheet_name or sheet_name not in self.current_sheets:
            return []
        
        try:
            df = self.current_sheets[sheet_name]
            non_empty_columns = []
            
            for col in df.columns:
                # Check if column has at least one non-empty value
                if not df[col].isna().all() and not (df[col].astype(str).str.strip() == '').all():
                    non_empty_columns.append(col)
            
            return non_empty_columns
        except Exception as e:
            print(f"Error getting non-empty columns: {str(e)}")
            return []
        
    def stack_values_merge(self, sheet_name, columns, new_column_name, delete_source=True):
        """
        Merge columns by stacking their values in separate rows.
        This preserves all data by creating new rows when needed.
        
        Args:
            sheet_name: Name of the sheet to modify
            columns: List of column names to merge
            new_column_name: Name for the merged column
            delete_source: Whether to delete source columns after merging
            
        Returns:
            bool: Success or failure
        """
        try:
            if sheet_name not in self.current_sheets:
                return False
                
            # Get the dataframe
            df = self.current_sheets[sheet_name].copy()
            
            # Check if all columns exist
            if not all(col in df.columns for col in columns):
                return False
            
            # Create a new dataframe to hold the stacked values
            new_rows = []
            
            # Track the original row indices for each value we stack
            for idx, row in df.iterrows():
                has_data = False
                # Create a new row with all columns except the ones we're merging
                base_row = {c: row[c] for c in df.columns if c not in columns}
                
                for col in columns:
                    if pd.notna(row[col]) and str(row[col]).strip() != "":
                        # Create a copy of the base row
                        new_row = base_row.copy()
                        # Add the value to the new column
                        new_row[new_column_name] = row[col]
                        new_rows.append(new_row)
                        has_data = True
                
                # If there was no data in any of the merged columns, add a row with empty value
                if not has_data:
                    base_row[new_column_name] = None
                    new_rows.append(base_row)
            
            # Create the new dataframe with stacked values
            if new_rows:
                stacked_df = pd.DataFrame(new_rows)
                
                # Add any missing columns from the original dataframe (should be rare)
                for col in df.columns:
                    if col not in stacked_df.columns and col not in columns:
                        stacked_df[col] = None
                
                # Apply the changes to the sheet
                self.current_sheets[sheet_name] = stacked_df
                
                # Delete source columns if needed
                if delete_source:
                    self.current_sheets[sheet_name] = self.current_sheets[sheet_name].drop(columns=columns, errors='ignore')
                
                # Mark that the sheet was modified
                self.modified_sheets.add(sheet_name)
                
                return True
            else:
                return False
                    
        except Exception as e:
            print(f"Error in stack_values_merge: {str(e)}")
            return False