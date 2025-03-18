import pandas as pd
import json
import os
from flask import Flask, jsonify

def fetch_dialfire_data():
    """Simulates fetching Dialfire call logs from an API."""
    file_path = os.path.join(os.path.dirname(__file__), "..", "data", "dialfire_call_logs.csv")
    file_path = os.path.abspath(file_path)  # Ensure absolute path resolution

    if not os.path.exists(file_path):
        return {"error": "Dialfire data not found"}
    
    df = pd.read_csv(file_path, index_col=0)
    return df.to_dict(orient="records")

if __name__ == "__main__":
    # Used when testing the function and printing results
    dialfire_data = fetch_dialfire_data()

    print("\nDialfire API Response:")
    print(json.dumps(dialfire_data, indent=4))