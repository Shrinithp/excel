from flask import Flask, request, send_file
from flask_cors import CORS  # Import CORS
import pandas as pd
import io
import os

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

@app.route('/export-excel', methods=['POST'])
def export_excel():
    try:
        data = request.json
        columns = data.get("columns", [])
        rows = data.get("data", [])
        sheet_name = data.get("sheet_name", "Sheet1")

        # Convert to DataFrame
        df = pd.DataFrame(rows, columns=columns)

        # Save to Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        output.seek(0)

        return send_file(output, as_attachment=True, download_name="exported_data.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
