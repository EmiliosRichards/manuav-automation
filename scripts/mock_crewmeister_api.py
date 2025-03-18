import os
import pandas as pd
import json
from flask import Flask, jsonify

app = Flask(__name__)

def fetch_crewmeister_data():
    """Simulates fetching Crewmeister work logs from an API."""
    file_path = os.path.join(os.path.dirname(__file__), "..", "data", "crewmeister_work_logs.xlsx")
    file_path = os.path.abspath(file_path)  # Ensure absolute path resolution
    
    if not os.path.exists(file_path):
        return {"error": "Crewmeister data not found"}
    
    df = pd.read_excel(file_path)
    return df.to_dict(orient="records")

@app.route("/crewmeister", methods=["GET"])
def get_crewmeister_data():
    """API endpoint to return Crewmeister work logs."""
    data = fetch_crewmeister_data()
    if "error" in data:
        return jsonify(data), 404  # Return HTTP 404 if data is missing
    return jsonify(data)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)

    # crewmeister_data = fetch_crewmeister_data()
    # print("Crewmeister API Response:")
    # print(json.dumps(crewmeister_data, indent=4))