import pandas as pd
import json
import os
from flask import Flask, jsonify

app = Flask(__name__)

def fetch_dialfire_data():
    """Simulates fetching Dialfire call logs from an API."""
    file_path = os.path.join(os.path.dirname(__file__), "..", "data", "dialfire_call_logs.csv")
    file_path = os.path.abspath(file_path)  # Ensure absolute path resolution

    if not os.path.exists(file_path):
        return {"error": "Dialfire data not found"}
    
    df = pd.read_csv(file_path, index_col=False)
    return df.to_dict(orient="records")

@app.route("/dialfire", methods=["GET"])
def get_dialfire_data():
    """API endpoint to return Dialfire call logs."""
    data = fetch_dialfire_data()
    if "error" in data:
        return jsonify(data), 404  # Return HTTP 404 if data is missing
    return jsonify(data)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)


# Used when testing the function and printing results
    # dialfire_data = fetch_dialfire_data()

    # print("\nDialfire API Response:")
    # print(json.dumps(dialfire_data, indent=4))

