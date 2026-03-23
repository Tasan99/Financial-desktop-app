from flask import Flask, jsonify, request
from flask_cors import CORS
from engine import FinancialDataEngine
import os

app = Flask(__name__)
CORS(app) 

# Motor başlangıçta boş (None) olacak
engine = None

# Yeni Endpoint: Frontend'den gelen dosya yolunu alıp motoru başlatır
@app.route('/api/load_file', methods=['POST'])
def load_file():
    global engine
    data = request.get_json()
    filepath = data.get('filepath')
    
    if not filepath or not os.path.exists(filepath):
        return jsonify({"error": "Invalid file path."}), 400
    
    try:
        # Motoru seçilen dosya ile başlat
        engine = FinancialDataEngine(filepath)
        return jsonify({"message": "System Ready", "entities": engine.ALL_ENTITIES})
    except Exception as e:
        return jsonify({"error": f"Excel Error: {str(e)}"}), 500

# Diğer endpoint'lerde motorun yüklü olup olmadığını kontrol et
@app.route('/api/entities', methods=['GET'])
def get_entities():
    if not engine: return jsonify([]), 400
    return jsonify(engine.ALL_ENTITIES)

@app.route('/api/report1', methods=['GET'])
def get_report1():
    if not engine: return jsonify({"error": "No data available"}), 400
    entity = request.args.get('entity', 'All Branches')
    return engine.build_report1_is_comparison(entity).to_json(orient='records')

@app.route('/api/report2', methods=['GET'])
def get_report2():
    if not engine: return jsonify({"error": "No data available"}), 400
    entity = request.args.get('entity', 'All Branches')
    period = request.args.get('period', 'Full Year Monthly')
    return engine.build_report2_trended_is(entity, period).to_json(orient='records')

@app.route('/api/report3', methods=['GET'])
def get_report3():
    if not engine: return jsonify({"error": "No data available"}), 400
    return engine.build_report3_rankings().to_json(orient='records')

@app.route('/api/report4', methods=['GET'])
def get_report4():
    if not engine: return jsonify({"error": "No data available"}), 400
    return engine.build_report4_regional().to_json(orient='records')

@app.route('/api/report5', methods=['GET'])
def get_report5():
    if not engine: return jsonify({"error": "No data available"}), 400
    entity = request.args.get('entity', 'All Branches')
    return engine.build_report5_metrics(entity).to_json(orient='records')

if __name__ == '__main__':
    app.run(port=5000, debug=False)