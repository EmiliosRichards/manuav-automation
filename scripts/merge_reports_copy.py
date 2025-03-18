import pandas as pd
import json
import os
import requests

# API Endpoints
CREWMEISTER_API = "http://localhost:5001/crewmeister"
DIALFIRE_API = "http://localhost:5000/dialfire"

def fetch_crewmeister_data():
    """Fetches Crewmeister work logs from the API."""
    response = requests.get(CREWMEISTER_API)
    return response.json() if response.status_code == 200 else {"error": "Failed to fetch Crewmeister data"}

def fetch_dialfire_data():
    """Fetches Dialfire call logs from the API."""
    response = requests.get(DIALFIRE_API)
    return response.json() if response.status_code == 200 else {"error": "Failed to fetch Dialfire data"}

def merge_and_clean_data():
    """Merges and cleans Crewmeister and Dialfire data."""
    crewmeister_data = fetch_crewmeister_data()
    dialfire_data = fetch_dialfire_data()

    # Handle API errors
    if isinstance(crewmeister_data, dict) and "error" in crewmeister_data:
        print(f"❌ Error: {crewmeister_data['error']}")
        return
    if isinstance(dialfire_data, dict) and "error" in dialfire_data:
        print(f"❌ Error: {dialfire_data['error']}")
        return

    
    # Convert JSON responses to DataFrames
    crew_df = pd.DataFrame(crewmeister_data)
    dialfire_df = pd.DataFrame(dialfire_data)

    print("Dialfire Columns:", dialfire_df.columns)  # Debugging step

    # Standardize column names (strip spaces, fix naming issues)
    crew_df.columns = crew_df.columns.str.strip()
    dialfire_df.columns = dialfire_df.columns.str.strip()

    # Standardizing agent names for merging
    name_mapping = {
        "Anna.Jager": "Anna Jager",
        "Frederike.Tietz.": "Frederike Tietz",
        "Lukas.Schneider": "Lukas Schneider",
        "Sven.Petro": "Sven Petro"
    }
    if "Agent Name" in dialfire_df.columns:
        dialfire_df["Agent Name"] = dialfire_df["Agent Name"].replace(name_mapping)

    # Convert date to standard format
    crew_df["Date"] = pd.to_datetime(crew_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    dialfire_df["Date"] = pd.to_datetime(dialfire_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

    # Identify missing values before replacing them
    missing_mask = merged_df.isnull()

    # Fill missing values with 0 to prevent calculation errors
    crew_df.fillna(0, inplace=True)
    dialfire_df.fillna(0, inplace=True)


    # Merging DataFrames on Agent Name and Date
    merged_df = pd.merge(crew_df, dialfire_df, on=["Agent Name", "Date"], how="outer")

    # Save the cleaned data to Excel
    os.makedirs("reports", exist_ok=True)  # Ensure reports folder exists

        # Highlight missing data
    for column in merged_df.columns:
        if merged_df[column].isnull().any():
            print(f"Warning: Missing values found in {column}")
    
    # Save to Excel with formatting
    report_path = "reports/Merged_Automation_Report.xlsx"
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        merged_df.to_excel(writer, sheet_name="Merged Data", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Merged Data"]
        
        # Apply formatting to highlight missing values
        format_highlight = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        for col_num, col_name in enumerate(merged_df.columns):
            worksheet.conditional_format(1, col_num, len(merged_df), col_num, {"type": "blanks", "format": format_highlight})
            missing_rows = missing_mask[col_name].values  # Find missing rows
            for row_num, is_missing in enumerate(missing_rows):
                if is_missing:
                    worksheet.write(row_num + 1, col_num, "", format_highlight)  # Highlight missing cell
    
    
    print(f"✅ Automation Report Generated Successfully! Saved to {report_path}")

if __name__ == "__main__":
    merge_and_clean_data()