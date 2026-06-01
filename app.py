import os
import sys
from flask import Flask, request, jsonify, render_template, send_file

project_dir = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=os.path.join(project_dir, 'templates'))
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB limit

@app.after_request
def add_cors_headers(response):
    """Injects CORS headers to permit cross-origin requests from local files (file:/// protocol)."""
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
    response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
    return response

@app.route('/')
def cable_routing_navigator():
    """Serves the Cable Routing Decision Navigator as the main landing page."""
    return render_template('cable_routing.html')

@app.route('/api/cable-routing/export-excel', methods=['POST', 'OPTIONS'])
def export_cable_routing_excel():
    """Populates routing inputs into the Excel template and returns the updated sheet stream."""
    if request.method == 'OPTIONS':
        return '', 200
        
    from io import BytesIO
    from openpyxl import load_workbook
    
    try:
        data = request.get_json() or {}
        inputs = data.get('inputs', {})
        
        template_path = os.path.join(project_dir, "Cable_Routing_Tool_Full_Logic.xlsx")
        if not os.path.exists(template_path):
            return jsonify({"success": False, "error": "Excel template not found."}), 404
            
        # Load template with formulas preserved
        wb = load_workbook(template_path, data_only=False)
        ws = wb['Routing Decision Tool']
        
        # Populate inputs (Col B)
        ws['B4'] = inputs.get('area', 'Outdoor')
        ws['B5'] = inputs.get('primary', 'FO Cable')
        ws['B6'] = inputs.get('shared', 'Power 400/230V')
        ws['B7'] = inputs.get('same', 'Yes')
        ws['B8'] = inputs.get('sep', 'Yes')
        ws['B9'] = inputs.get('sched', 'Same Phase')
        try:
            ws['B10'] = float(inputs.get('spare', 20)) / 100.0
        except (ValueError, TypeError):
            ws['B10'] = 0.20
            
        ws['B11'] = inputs.get('ownerApprove', 'Yes')
            
        # Output cells (B13-B22) are left unmodified to preserve the template's Excel formulas.
        # Excel will dynamically calculate these cells based on the inputs when opened.
        
        # Populate notes in A25 (replaces placeholder text)
        if inputs.get('notes'):
            ws['A25'] = inputs.get('notes')
            
        # Stream file back to client
        out = BytesIO()
        wb.save(out)
        out.seek(0)
        
        return send_file(
            out,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="Cable_Routing_Decision_Output.xlsx"
        )
    except Exception as e:
        print(f"Excel export error: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    print("Starting Standalone Cable Routing Decision Navigator at http://127.0.0.1:5000...")
    app.run(host='127.0.0.1', port=5000, debug=False)
