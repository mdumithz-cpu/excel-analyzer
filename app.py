# Excel Analyzer - Web Application (Flask)
# File: app.py

from flask import Flask, render_template, request, send_file, jsonify, flash, redirect, url_for
import pandas as pd
import numpy as np
import re
import os
import io
from datetime import datetime
from werkzeug.utils import secure_filename
import tempfile
import uuid

app = Flask(__name__)
app.secret_key = 'excel_analyzer_secret_key_2024'  # Change this in production
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
    """Excel processing logic (same as desktop app)"""
    
    @staticmethod
    def load_excel_file(file_path):
        """Load and prepare Excel file"""
        try:
            # Read Excel file using row 5 (6th row) as header
            df = pd.read_excel(file_path, header=5)
            
            # Clean column names
            cleaned_columns = []
            for i, col in enumerate(df.columns):
                if pd.isna(col) or str(col).strip() == '' or str(col).startswith('Unnamed'):
                    cleaned_columns.append(f"Col_{i}")
                else:
                    cleaned_columns.append(str(col).strip())
            
            df.columns = cleaned_columns
            
            # Remove empty rows
            df = df.dropna(how='all').reset_index(drop=True)
            
            return df, None
            
        except Exception as e:
            return None, str(e)
    
    @staticmethod
    def filter_and_process_data(df):
        """Filter and process data according to criteria"""
        try:
            # Find name column
            name_col = None
            for col in df.columns:
                if col.lower() == 'name':
                    name_col = col
                    break
            
            if name_col is None:
                return None, f"'name' column not found. Available columns: {list(df.columns)}"
            
            # Filter rows
            name_criteria = ["Input rate alarm notification", "Output rate alarm notification"]
            filtered_df = df[df[name_col].isin(name_criteria)].copy()
            
            if len(filtered_df) == 0:
                return None, "No rows match the name criteria!"
            
            # Find required columns
            required_columns = ["name", "Alarm Source", "Location Info", "Arrived On (ST)", "Other Information"]
            column_mapping = {}
            
            for req_col in required_columns:
                found_col = ExcelProcessor.find_column(filtered_df.columns, req_col)
                if found_col:
                    column_mapping[req_col] = found_col
            
            if not column_mapping:
                return None, "No required columns found!"
            
            # Select and rename columns
            available_cols = [column_mapping[col] for col in required_columns if col in column_mapping]
            filtered_df = filtered_df[available_cols].copy()
            rename_dict = {v: k for k, v in column_mapping.items()}
            filtered_df = filtered_df.rename(columns=rename_dict)
            
            # Handle duplicates
            if all(col in filtered_df.columns for col in ['Alarm Source', 'Location Info', 'Arrived On (ST)']):
                filtered_df['Arrived On (ST)'] = pd.to_datetime(filtered_df['Arrived On (ST)'], errors='coerce')
                filtered_df = filtered_df.sort_values('Arrived On (ST)')
                filtered_df = filtered_df.drop_duplicates(subset=['Alarm Source', 'Location Info'], keep='first')
            
            return filtered_df, None
            
        except Exception as e:
            return None, str(e)
    
    @staticmethod
    def find_column(columns, target):
        """Find column with flexible matching"""
        # Exact match
        for col in columns:
            if col.lower().strip() == target.lower().strip():
                return col
        
        # Partial match
        for col in columns:
            if target.lower().replace(" ", "") in col.lower().replace(" ", ""):
                return col
        
        # Variation matching
        variations = {
            "name": ["name", "alarm name", "description"],
            "alarm source": ["alarm source", "source", "alarm_source"],
            "location info": ["location info", "location", "location_info"],
            "arrived on (st)": ["arrived on", "arrived", "time", "timestamp"],
            "other information": ["other information", "other info", "other", "details"]
        }
        
        if target.lower() in variations:
            for variant in variations[target.lower()]:
                for col in columns:
                    if variant in col.lower():
                        return col
        
        return None
    
    @staticmethod
    def extract_node_b(location_info):
        """Extract Node B from Location Info"""
        if pd.isna(location_info):
            return ""
        
        location_str = str(location_info)
        node_b = ""
        
        # Look for "To" or "to" patterns
        patterns = [
            r'(?:To|to)\s+([^=]+?)(?=\s*=|$)',
            r'TO\s+([^=]+?)(?=\s*=|$)',
            r'LINK[-\s]*TO\s+([^_\s]+(?:_[^_\s]+)*)',
            r'---\s*(.+?)(?=\s*\(|$)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, location_str, re.IGNORECASE)
            if match:
                node_b = match.group(1).strip()
                break
        
        # Clean Node B
        if node_b:
            cleanup_pattern = r'[-]{1,}|[#]{1,}|\('
            match = re.search(cleanup_pattern, node_b)
            if match:
                node_b = node_b[:match.start()].strip()
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
        
        return ""
    
    @staticmethod
    def extract_utilization_percentage(other_info):
        """Extract percentage value from Other Information"""
        if pd.isna(other_info):
            return ""
        
        other_str = str(other_info)
        
        # Standard percentage pattern
        pattern1 = r'(\d+(?:\.\d+)?)\s*%'
        match = re.search(pattern1, other_str)
        if match:
            return f"{match.group(1)}%"
        
        # Context-aware extraction
        pattern2 = r'(?:utilization|usage|load)[\s:]*(\d+(?:\.\d+)?)'
        match = re.search(pattern2, other_str, re.IGNORECASE)
        if match:
            return f"{match.group(1)}%"
        
        return ""
    
    @staticmethod
    def create_output_table(filtered_df):
        """Create the final output table"""
        try:
            # Create output DataFrame
            output_df = pd.DataFrame()
            
            # Add columns
            output_df['Index'] = range(1, len(filtered_df) + 1)
            
            if 'Arrived On (ST)' in filtered_df.columns:
                output_df['Occurred Time'] = filtered_df['Arrived On (ST)'].reset_index(drop=True)
            else:
                output_df['Occurred Time'] = ""
            
            if 'Alarm Source' in filtered_df.columns:
                output_df['Node A'] = filtered_df['Alarm Source'].reset_index(drop=True)
            else:
                output_df['Node A'] = ""
            
            if 'Location Info' in filtered_df.columns:
                output_df['Node B'] = filtered_df['Location Info'].apply(ExcelProcessor.extract_node_b).reset_index(drop=True)
                output_df['Link Description'] = filtered_df['Location Info'].apply(ExcelProcessor.extract_link_description).reset_index(drop=True)
            else:
                output_df['Node B'] = ""
                output_df['Link Description'] = ""
            
            if 'Other Information' in filtered_df.columns:
                output_df['Utilization%'] = filtered_df['Other Information'].apply(ExcelProcessor.extract_utilization_percentage).reset_index(drop=True)
            else:
                output_df['Utilization%'] = ""
            
            # Add empty columns
            output_df['Reason'] = ""
            output_df['Remarks'] = ""
            
            # Remove duplicates based on Node A, Node B, Link Description
            output_df = output_df.drop_duplicates(subset=['Node A', 'Node B', 'Link Description'], keep='first')
            
            # Advanced S pattern duplicate removal
            output_df = ExcelProcessor.remove_s_pattern_duplicates(output_df)
            
            # Group by Node A and sort
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
            return match.group(0) if match else None
        
        output_df['_temp_s_pattern'] = output_df['Link Description'].apply(extract_s_pattern)
        s_pattern_rows = output_df[output_df['_temp_s_pattern'].notna()].copy()
        
        if len(s_pattern_rows) > 0:
            s_pattern_counts = s_pattern_rows['_temp_s_pattern'].value_counts()
            duplicate_s_patterns = s_pattern_counts[s_pattern_counts > 1].index.tolist()
            
            if duplicate_s_patterns:
                rows_to_remove = []
                for s_pattern in duplicate_s_patterns:
                    pattern_rows = output_df[output_df['_temp_s_pattern'] == s_pattern].copy()
                    pattern_rows = pattern_rows.sort_values('Occurred Time')
                    
                    if len(pattern_rows) > 1:
                        rows_to_remove.extend(pattern_rows.iloc[1:].index.tolist())
                
                if rows_to_remove:
                    output_df = output_df.drop(rows_to_remove)
        
        output_df = output_df.drop('_temp_s_pattern', axis=1)
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
            # Generate unique filename
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4()}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            
            # Save uploaded file
            file.save(filepath)
            
            # Process the Excel file
            processor = ExcelProcessor()
            
            # Step 1: Load Excel file
            df, error = processor.load_excel_file(filepath)
            if error:
                os.remove(filepath)  # Clean up
                return jsonify({'success': False, 'error': f'Error loading file: {error}'})
            
            # Step 2: Filter and process data
            filtered_data, error = processor.filter_and_process_data(df)
            if error:
                os.remove(filepath)  # Clean up
                return jsonify({'success': False, 'error': f'Error processing data: {error}'})
            
            # Step 3: Create output table
            output_table, error = processor.create_output_table(filtered_data)
            if error:
                os.remove(filepath)  # Clean up
                return jsonify({'success': False, 'error': f'Error creating output: {error}'})
            
            # Step 4: Save processed file
            output_filename = f"processed_{unique_filename}"
            output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            
            output_table.to_excel(output_filepath, index=False)
            
            # Clean up input file
            os.remove(filepath)
            
            # Return success with download link
            return jsonify({
                'success': True, 
                'message': f'File processed successfully! {len(output_table)} rows generated.',
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
    """Clean up old files (run periodically)"""
    try:
        current_time = datetime.now().timestamp()
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file_time = os.path.getctime(filepath)
            
            # Delete files older than 1 hour
            if current_time - file_time > 3600:
                os.remove(filepath)
        
        return "Cleanup completed"
    except Exception as e:
        return f"Cleanup error: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))