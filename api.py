# API with endpoint to upload excel file

from flask import Flask, request, jsonify
from flask_cors import CORS
from libs.excel_parser import ExcelParser

app = Flask(__name__)
CORS(app)

@app.route('/parse', methods=['POST'])
def parse():
    excel_file = request.files['file']
    excel_parser = ExcelParser()
    result_json = excel_parser.parse_xlsx_to_json_file(excel_file)
    if result_json:
        return result_json
    else:
        return jsonify({"error": "Error parsing file"})

if __name__ == '__main__':
    app.run(debug=True)