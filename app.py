# Excel Analyzer - Web Application (Flask)
# File: app.py

from flask import Flask, render_template, request, send_file, jsonify, url_for
import pandas as pd
import numpy as np
import re
import os
from datetime import datetime
from werkzeug.utils import secure_filename
import uuid

app = Flask(__name__)
app.secret_key = 'excel_analyzer_secret_key_2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload directory exists
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

class ExcelProcessor:
    """Excel processing logic - complete implementation from Colab script"""
    
    @staticmethod
    def load_excel_file(file_path):
        """Load and prepare Excel file"""
        try:
            # Read Excel file using row 5 (6th row) as header, which automatically skips first 5 rows
            df = pd.read_excel(file_path, header=5)
            
            # Clean column names - remove extra whitespace and handle unnamed/NaN columns
            cleaned_columns = []
            for i, col in enumerate(df.columns):
                if pd.isna(col) or str(col).strip() == '' or str(col).startswith('Unnamed'):
                    cleaned_columns.append(f"Col_{i}")
                else:
                    cleaned_columns.append(str(col).strip())
            
            df.columns = cleaned_columns
            
            # Remove any completely empty rows
            df = df.dropna(how='all').reset_index(drop=True)
            
            return df, None
            
        except Exception as e:
            return None, str(e)
    
    @staticmethod
    def filter_and_process_data(df):
        """Filter and process data according to specified criteria"""
        try:
            # Filter rows with specific names - UPDATED TO INCLUDE "Performance Critical Alarm"
            name_criteria = [
                "Input rate alarm notification", 
                "Output rate alarm notification",
                "Performance Critical Alarm"
            ]

            # Check if 'name' column exists (case-insensitive)
            name_col = None
            for col in df.columns:
                if col.lower() == 'name':
                    name_col = col
                    break

            if name_col is None:
                return None, f"'name' column not found. Available columns: {list(df.columns)}"

            # Filter rows
            filtered_df = df[df[name_col].isin(name_criteria)].copy()

            if len(filtered_df) == 0:
                return None, "No rows match the name criteria!"

            # Select required columns (case-insensitive matching)
            required_columns = ["name", "Alarm Source", "Location Info", "Arrived On (ST)", "Other Information"]
            column_mapping = {}

            for req_col in required_columns:
                found_col = ExcelProcessor.find_column(filtered_df.columns, req_col)
                if found_col:
                    column_mapping[req_col] = found_col

            if not column_mapping:
                return None, "No required columns found!"

            # Select only the columns that were found
            available_cols = [column_mapping[col] for col in required_columns if col in column_mapping]
            filtered_df = filtered_df[available_cols].copy()

            # Rename columns to standard names
            rename_dict = {v: k for k, v in column_mapping.items()}
            filtered_df = filtered_df.rename(columns=rename_dict)

            # Convert "Arrived On (ST)" to datetime
            if 'Arrived On (ST)' in filtered_df.columns:
                filtered_df['Arrived On (ST)'] = pd.to_datetime(filtered_df['Arrived On (ST)'], errors='coerce')

            # Extract utilization percentage for deduplication - UPDATED LOGIC
            if 'Other Information' in filtered_df.columns and 'name' in filtered_df.columns:
                filtered_df['_temp_utilization'] = filtered_df.apply(
                    lambda row: ExcelProcessor.extract_utilization_percentage(
                        row['Other Information'], 
                        row['name']
                    ), 
                    axis=1
                )
                # Convert to numeric for comparison (remove % sign)
                filtered_df['_temp_util_numeric'] = filtered_df['_temp_utilization'].apply(
                    lambda x: float(x.replace('%', '')) if x and x != '' else 0.0
                )
            else:
                filtered_df['_temp_util_numeric'] = 0.0

            # Handle duplicates - keep ONLY HIGHEST utilization (no timestamp dependency)
            # This works across ALL alarm types including Performance Critical Alarm and Input/Output rate alarms
            if all(col in filtered_df.columns for col in ['Alarm Source', 'Location Info']):
                # Sort by utilization (descending only)
                filtered_df = filtered_df.sort_values('_temp_util_numeric', ascending=False)
                # Drop duplicates keeping first occurrence (highest utilization only)
                # This automatically handles mixed alarm types since we compare numeric utilization values
                filtered_df = filtered_df.drop_duplicates(subset=['Alarm Source', 'Location Info'], keep='first')
            
            # Remove temporary columns
            filtered_df = filtered_df.drop(['_temp_utilization', '_temp_util_numeric'], axis=1, errors='ignore')
            
            return filtered_df, None
            
        except Exception as e:
            return None, str(e)
    
    @staticmethod
    def find_column(columns, target):
        """Find column with flexible matching"""
        # Try exact match first
        for col in columns:
            if col.lower().strip() == target.lower().strip():
                return col
        
        # If not found, try partial matching
        for col in columns:
            if target.lower().replace(" ", "") in col.lower().replace(" ", ""):
                return col
        
        # Try even more flexible matching for common variations
        variations = {
            "name": ["name", "alarm name", "description", "alarm_name"],
            "alarm source": ["alarm source", "source", "alarm_source", "alarmsource"],
            "location info": ["location info", "location", "location_info", "locationinfo"],
            "arrived on (st)": ["arrived on", "arrived", "time", "timestamp", "arrived_on", "date"],
            "other information": ["other information", "other info", "other", "details", "other_information"]
        }
        
        if target.lower() in variations:
            for variant in variations[target.lower()]:
                for col in columns:
                    if variant in col.lower():
                        return col
        
        return None
    
    @staticmethod
    def contains_f8(node_value):
        """Check if Node A or Node B contains F8 pattern (e.g., xx_F8_CEA or xx_F8)"""
        if pd.isna(node_value) or node_value == "":
            return False
        
        node_str = str(node_value).upper()
        
        # Pattern 1: F8 with underscore before/after (e.g., xx_F8_CEA, xx_F8, F8_xx)
        pattern1 = r'_F8_|_F8$|^F8_'
        if re.search(pattern1, node_str):
            return True
        
        # Pattern 2: F8 as standalone or with non-alphanumeric boundaries
        pattern2 = r'\bF8\b'
        if re.search(pattern2, node_str):
            return True
        
        return False
    
    @staticmethod
    def extract_node_b(location_info):
        """Extract Node B from Location Info"""
        if pd.isna(location_info):
            return ""

        location_str = str(location_info)
        node_b = ""

        # Pattern 1: Look for "To" or "to" followed by text until "=" or end
        pattern1 = r'(?:To|to)\s+([^=]+?)(?=\s*=|$)'
        match = re.search(pattern1, location_str)

        if match:
            node_b = match.group(1).strip()
        else:
            # Pattern 2: Look for "TO" in uppercase
            pattern2 = r'TO\s+([^=]+?)(?=\s*=|$)'
            match = re.search(pattern2, location_str)

            if match:
                node_b = match.group(1).strip()
            else:
                # Pattern 3: If no "To" found, look for common network node patterns
                pattern3 = r'LINK[-\s]*TO\s+([^_\s]+(?:_[^_\s]+)*)'
                match = re.search(pattern3, location_str, re.IGNORECASE)

                if match:
                    node_b = match.group(1).strip()
                else:
                    # Pattern 4: Extract anything after "---" or similar separators
                    pattern4 = r'---\s*(.+?)(?=\s*\(|$)'
                    match = re.search(pattern4, location_str)

                    if match:
                        node_b = match.group(1).strip()

        # Clean the extracted Node B
        if node_b:
            # Remove text starting from dashes, hashes, or opening parenthesis
            cleanup_pattern = r'[-]{1,}|[#]{1,}|\('
            match = re.search(cleanup_pattern, node_b)
            if match:
                node_b = node_b[:match.start()].strip()

            # Final cleanup: remove trailing punctuation or whitespace
            node_b = re.sub(r'[,;\s_]+$', '', node_b)

        return node_b
    
    @staticmethod
    def extract_link_description(location_info):
        """Extract Link Description from Location Info"""
        if pd.isna(location_info):
            return ""

        location_str = str(location_info)

        # Pattern 1: Text between # symbols
        pattern1 = r'#([^#]+)#'
        match = re.search(pattern1, location_str)

        if match:
            return match.group(1).strip()

        # Pattern 2: Text between underscores
        pattern2 = r'_([^_]+)_'
        match = re.search(pattern2, location_str)

        if match:
            return match.group(1).strip()

        # Pattern 3: Extract main part of link names
        if '_' in location_str:
            parts = location_str.split('_')
            if len(parts) >= 3:
                return '_'.join(parts[1:-1])
            elif len(parts) == 2:
                return parts[1]

        # Pattern 4: Look for descriptive parts
        pattern4 = r'([A-Za-z][A-Za-z0-9_-]*[A-Za-z])'
        matches = re.findall(pattern4, location_str)

        if matches:
            longest_match = max(matches, key=len)
            if len(longest_match) > 3:
                return longest_match

        return ""
    
    @staticmethod
    def extract_utilization_percentage(other_info, alarm_name=None):
        """
        Extract utilization percentage based on alarm type
        
        For "Performance Critical Alarm": Extract from "Indicator Value=xxx.xx %"
        For other alarms: Extract HIGHEST percentage value between Input and Output flow bandwidth usage
        """
        if pd.isna(other_info):
            return ""
        
        other_str = str(other_info)
        
        # Check if this is a "Performance Critical Alarm"
        if alarm_name and "Performance Critical Alarm" in str(alarm_name):
            # Extract from "Indicator Value=xxx.xx %"
            indicator_pattern = r'Indicator\s+Value\s*=\s*(\d+(?:\.\d+)?)\s*%?'
            match = re.search(indicator_pattern, other_str, re.IGNORECASE)
            
            if match:
                percentage = float(match.group(1))
                if percentage % 1 == 0:
                    return f"{int(percentage)}%"
                else:
                    return f"{percentage}%"
            
            return ""
        
        # For other alarm types (Input/Output rate alarm notification)
        percentages = []
        
        # Pattern for Input flow bandwidth usage
        input_pattern = r'Input\s+flow\s+bandwidth\s+usage\s*=\s*(\d+(?:\.\d+)?)\s*%?'
        input_matches = re.findall(input_pattern, other_str, re.IGNORECASE)
        percentages.extend([float(m) for m in input_matches])
        
        # Pattern for Output flow bandwidth usage
        output_pattern = r'Output\s+flow\s+bandwidth\s+usage\s*=\s*(\d+(?:\.\d+)?)\s*%?'
        output_matches = re.findall(output_pattern, other_str, re.IGNORECASE)
        percentages.extend([float(m) for m in output_matches])
        
        # Return the HIGHEST percentage found
        if percentages:
            max_percentage = max(percentages)
            if max_percentage % 1 == 0:
                return f"{int(max_percentage)}%"
            else:
                return f"{max_percentage}%"
        
        return ""
    
    @staticmethod
    def create_output_table(filtered_df):
        """Create the final output table with required columns"""
        try:
            output_df = pd.DataFrame()

            # Index
            output_df['Index'] = range(1, len(filtered_df) + 1)

            # Occurred Time
            if 'Arrived On (ST)' in filtered_df.columns:
                output_df['Occurred Time'] = filtered_df['Arrived On (ST)'].reset_index(drop=True)
            else:
                output_df['Occurred Time'] = ""

            # Node A
            if 'Alarm Source' in filtered_df.columns:
                output_df['Node A'] = filtered_df['Alarm Source'].reset_index(drop=True)
            else:
                output_df['Node A'] = ""

            # Node B
            if 'Location Info' in filtered_df.columns:
                output_df['Node B'] = filtered_df['Location Info'].apply(ExcelProcessor.extract_node_b).reset_index(drop=True)
            else:
                output_df['Node B'] = ""

            # Link Description
            if 'Location Info' in filtered_df.columns:
                output_df['Link Description'] = filtered_df['Location Info'].apply(ExcelProcessor.extract_link_description).reset_index(drop=True)
            else:
                output_df['Link Description'] = ""

            # Utilization% - UPDATED TO PASS ALARM NAME
            if 'Other Information' in filtered_df.columns and 'name' in filtered_df.columns:
                output_df['Utilization%'] = filtered_df.apply(
                    lambda row: ExcelProcessor.extract_utilization_percentage(
                        row['Other Information'], 
                        row['name']
                    ), 
                    axis=1
                ).reset_index(drop=True)
            elif 'Other Information' in filtered_df.columns:
                output_df['Utilization%'] = filtered_df['Other Information'].apply(
                    lambda x: ExcelProcessor.extract_utilization_percentage(x)
                ).reset_index(drop=True)
            else:
                output_df['Utilization%'] = ""

            # Reason and Remarks
            output_df['Reason'] = ""
            output_df['Remarks'] = ""

            # Create numeric utilization for comparison
            output_df['_temp_util_numeric'] = output_df['Utilization%'].apply(
                lambda x: float(x.replace('%', '')) if x and x != '' else 0.0
            )

            # Remove duplicates - keep HIGHEST utilization only
            # This works across ALL alarm types (Performance Critical Alarm, Input/Output rate alarms)
            # For the same Node A, Node B & Link Description combination, we keep the row with max utilization
            # regardless of whether it comes from "Indicator Value" or "Input/Output flow bandwidth usage"
            if len(output_df) > 0:
                output_df = output_df.sort_values('_temp_util_numeric', ascending=False)
                output_df = output_df.drop_duplicates(
                    subset=['Node A', 'Node B', 'Link Description'], 
                    keep='first'
                )

            # Remove S-pattern duplicates
            output_df = ExcelProcessor.remove_s_pattern_duplicates(output_df)

            # Remove temporary column
            output_df = output_df.drop('_temp_util_numeric', axis=1, errors='ignore')

            # Sort by Node A and Occurred Time
            if len(output_df) > 0:
                output_df = output_df.sort_values(['Node A', 'Occurred Time']).reset_index(drop=True)
                output_df['Index'] = range(1, len(output_df) + 1)
            
            return output_df, None
            
        except Exception as e:
            return None, str(e)
    
    @staticmethod
    def remove_s_pattern_duplicates(output_df):
        """Remove duplicates based on S pattern in Link Description"""
        def extract_s_pattern(link_desc):
            if pd.isna(link_desc) or link_desc == "":
                return None
            pattern = r'S(\d+)(?:_\d+)?'
            match = re.search(pattern, str(link_desc))
            if match:
                return match.group(0)
            return None

        output_df['_temp_s_pattern'] = output_df['Link Description'].apply(extract_s_pattern)

        if '_temp_util_numeric' not in output_df.columns:
            output_df['_temp_util_numeric'] = output_df['Utilization%'].apply(
                lambda x: float(x.replace('%', '')) if x and x != '' else 0.0
            )

        s_pattern_rows = output_df[output_df['_temp_s_pattern'].notna()].copy()

        if len(s_pattern_rows) > 0:
            s_pattern_counts = s_pattern_rows['_temp_s_pattern'].value_counts()
            duplicate_s_patterns = s_pattern_counts[s_pattern_counts > 1].index.tolist()

            if duplicate_s_patterns:
                rows_to_remove = []

                for s_pattern in duplicate_s_patterns:
                    pattern_rows = output_df[output_df['_temp_s_pattern'] == s_pattern].copy()
                    pattern_rows = pattern_rows.sort_values('_temp_util_numeric', ascending=False)

                    if len(pattern_rows) > 1:
                        rows_to_drop = pattern_rows.iloc[1:]
                        rows_to_remove.extend(rows_to_drop.index.tolist())

                if rows_to_remove:
                    output_df = output_df.drop(rows_to_remove)

        output_df = output_df.drop('_temp_s_pattern', axis=1, errors='ignore')
        return output_df

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file selected'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No file selected'})
    
    if file and allowed_file(file.filename):
        try:
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4()}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            
            file.save(filepath)
            
            shift_in_date = request.form.get('shift_in_date', '')
            shift_in_time = request.form.get('shift_in_time', '')
            shift_out_date = request.form.get('shift_out_date', '')
            shift_out_time = request.form.get('shift_out_time', '')
            
            processor = ExcelProcessor()
            
            df, error = processor.load_excel_file(filepath)
            if error:
                os.remove(filepath)
                return jsonify({'success': False, 'error': f'Error loading file: {error}'})
            
            filtered_data, error = processor.filter_and_process_data(df)
            if error:
                os.remove(filepath)
                return jsonify({'success': False, 'error': f'Error processing data: {error}'})
            
            if shift_in_date and shift_in_time and shift_out_date and shift_out_time:
                try:
                    shift_in_datetime = pd.to_datetime(f"{shift_in_date} {shift_in_time}")
                    shift_out_datetime = pd.to_datetime(f"{shift_out_date} {shift_out_time}")
                    
                    if 'Arrived On (ST)' in filtered_data.columns:
                        filtered_data['Arrived On (ST)'] = pd.to_datetime(filtered_data['Arrived On (ST)'], errors='coerce')
                        
                        filtered_data = filtered_data[
                            (filtered_data['Arrived On (ST)'] >= shift_in_datetime) & 
                            (filtered_data['Arrived On (ST)'] <= shift_out_datetime)
                        ]
                        
                        if len(filtered_data) == 0:
                            os.remove(filepath)
                            return jsonify({'success': False, 'error': f'No alarms found between {shift_in_datetime} and {shift_out_datetime}'})
                        
                except Exception as e:
                    os.remove(filepath)
                    return jsonify({'success': False, 'error': f'Error applying date filter: {str(e)}'})
            
            output_table, error = processor.create_output_table(filtered_data)
            if error:
                os.remove(filepath)
                return jsonify({'success': False, 'error': f'Error creating output: {error}'})
            
            output_filename = f"processed_{unique_filename}"
            output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            
            output_table.to_excel(output_filepath, index=False)
            
            os.remove(filepath)
            
            message = f'File processed successfully! {len(output_table)} rows generated.'
            if shift_in_date and shift_out_date:
                message += f' Filtered from {shift_in_date} {shift_in_time} to {shift_out_date} {shift_out_time}.'
            
            return jsonify({
                'success': True, 
                'message': message,
                'download_url': url_for('download_file', filename=output_filename),
                'filename': output_filename,
                'stats': {
                    'total_rows': len(output_table),
                    'node_b_count': len(output_table[output_table['Node B'] != '']),
                    'link_desc_count': len(output_table[output_table['Link Description'] != '']),
                    'utilization_count': len(output_table[output_table['Utilization%'] != ''])
                }
            })
            
        except Exception as e:
            return jsonify({'success': False, 'error': f'Processing error: {str(e)}'})
    
    return jsonify({'success': False, 'error': 'Invalid file type. Please upload .xlsx or .xls files.'})

@app.route('/download/<filename>')
def download_file(filename):
    try:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True, download_name=f"excel_analysis_result.xlsx")
        else:
            return "File not found", 404
    except Exception as e:
        return f"Error downloading file: {str(e)}", 500

@app.route('/cleanup')
def cleanup_files():
    """Clean up old files"""
    try:
        current_time = datetime.now().timestamp()
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file_time = os.path.getctime(filepath)
            
            if current_time - file_time > 3600:
                os.remove(filepath)
        
        return "Cleanup completed"
    except Exception as e:
        return f"Cleanup error: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
