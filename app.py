from flask import Flask, jsonify, request
from flask_cors import CORS
import xlwings as xw
import pandas as pd
import json
import numpy as np
from datetime import datetime
from decimal import Decimal
from config import DEFAULT_TEST_CELLS

app = Flask(__name__)
CORS(app, origins=["https://localhost:3000"], methods=["GET", "POST", "OPTIONS"])

class RobustJSONEncoder(json.JSONEncoder):
    """Custom JSON encoder that handles numpy types, datetime, and decimal."""
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif isinstance(obj, (datetime, pd.Timestamp)):
            return obj.isoformat()
        elif isinstance(obj, Decimal):
            return float(obj)
        elif pd.isna(obj):
            return None
        return super().default(obj)

app.json_encoder = RobustJSONEncoder

def get_excel_app():
    """Get the active Excel application."""
    try:
        return xw.apps.active
    except Exception:
        return None

def get_active_workbook(app=None):
    """Get the active workbook."""
    if app is None:
        app = get_excel_app()
    if app is None:
        return None
    try:
        return app.books.active
    except Exception:
        return None

def get_worksheet(workbook, sheet_name=None):
    """Get a specific worksheet or the active one."""
    if sheet_name:
        try:
            return workbook.sheets[sheet_name]
        except Exception:
            return None
    try:
        return workbook.sheets.active
    except Exception:
        return None

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    try:
        excel_app = get_excel_app()
        if excel_app is None:
            return jsonify({"status": "unhealthy", "error": "No Excel application running"}), 503
        
        workbook = get_active_workbook(excel_app)
        if workbook is None:
            return jsonify({"status": "unhealthy", "error": "No active workbook found"}), 404
        
        return jsonify({"status": "healthy", "workbook": workbook.name})
    except Exception as e:
        return jsonify({"status": "unhealthy", "error": str(e)}), 500

@app.route('/api/excel-data', methods=['GET'])
def get_excel_data():
    """Get data from Excel with optional workbook/sheet targeting."""
    try:
        # Get query parameters
        workbook_name = request.args.get('workbook')
        sheet_name = request.args.get('sheet')
        include_cell_mapping = request.args.get('include_cell_mapping', 'true').lower() == 'true'
        
        excel_app = get_excel_app()
        if excel_app is None:
            return jsonify({"error": "No Excel application running"}), 503
        
        # Get workbook
        if workbook_name:
            try:
                workbook = excel_app.books[workbook_name]
            except Exception:
                return jsonify({"error": f"Workbook '{workbook_name}' not found"}), 404
        else:
            workbook = get_active_workbook(excel_app)
            if workbook is None:
                return jsonify({"error": "No active workbook found"}), 404
        
        # Get worksheet
        worksheet = get_worksheet(workbook, sheet_name)
        if worksheet is None:
            if sheet_name:
                return jsonify({"error": f"Sheet '{sheet_name}' not found in workbook '{workbook.name}'"}), 404
            else:
                return jsonify({"error": "No active worksheet found"}), 404
        
        # Get used range
        used_range = worksheet.used_range
        if used_range is None:
            return jsonify({"error": "No data found in worksheet"}), 404
        
        # Get data as DataFrame
        df = used_range.options(pd.DataFrame, header=1, index=False).value
        
        if df is None or df.empty:
            return jsonify({"error": "No data found in worksheet"}), 404
        
        # Prepare response
        response_data = {
            "workbook": workbook.name,
            "sheet": worksheet.name,
            "data": df.to_dict('records'),
            "shape": df.shape
        }
        
        # Add cell mapping for smaller datasets
        if include_cell_mapping and df.shape[0] <= 100:
            cell_mapping = {}
            start_row = used_range.row
            start_col = used_range.column
            
            for i, row in df.iterrows():
                for j, col in enumerate(df.columns):
                    cell_addr = xw.utils.int_to_col_letter(start_col + j) + str(start_row + i + 1)
                    cell_mapping[cell_addr] = row[col]
            
            response_data["cell_mapping"] = cell_mapping
        
        return jsonify(response_data)
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/write-excel', methods=['POST'])
def write_excel_data():
    """Write data to Excel with comprehensive targeting support."""
    try:
        data = request.get_json()
        if not data or 'operations' not in data:
            return jsonify({"error": "No operations provided"}), 400
        
        operations = data['operations']
        workbook_name = data.get('workbook')
        sheet_name = data.get('sheet')
        
        excel_app = get_excel_app()
        if excel_app is None:
            return jsonify({"error": "No Excel application running"}), 503
        
        # Get workbook
        if workbook_name:
            try:
                workbook = excel_app.books[workbook_name]
            except Exception:
                return jsonify({"error": f"Workbook '{workbook_name}' not found"}), 404
        else:
            workbook = get_active_workbook(excel_app)
            if workbook is None:
                return jsonify({"error": "No active workbook found"}), 404
        
        # Get worksheet
        worksheet = get_worksheet(workbook, sheet_name)
        if worksheet is None:
            if sheet_name:
                return jsonify({"error": f"Sheet '{sheet_name}' not found in workbook '{workbook.name}'"}), 404
            else:
                return jsonify({"error": "No active worksheet found"}), 404
        
        results = []
        
        for operation in operations:
            op_type = operation.get('type')
            
            if op_type == 'write_cell':
                cell = operation.get('cell')
                value = operation.get('value')
                
                if not cell:
                    results.append({"error": "Cell address required for write_cell operation"})
                    continue
                
                try:
                    worksheet.range(cell).value = value
                    results.append({"success": f"Written '{value}' to cell {cell}"})
                except Exception as e:
                    results.append({"error": f"Failed to write to cell {cell}: {str(e)}"})
            
            elif op_type == 'write_range':
                range_addr = operation.get('range')
                values = operation.get('values')
                
                if not range_addr or not values:
                    results.append({"error": "Range and values required for write_range operation"})
                    continue
                
                try:
                    worksheet.range(range_addr).value = values
                    results.append({"success": f"Written data to range {range_addr}"})
                except Exception as e:
                    results.append({"error": f"Failed to write to range {range_addr}: {str(e)}"})
            
            else:
                results.append({"error": f"Unknown operation type: {op_type}"})
        
        return jsonify({"results": results})
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=3001, debug=True)