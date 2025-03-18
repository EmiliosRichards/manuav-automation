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

# 📌 Generate a Human-Readable Error Summary
def generate_error_summary(merged_df):
    """Generates a detailed error summary with explanations and suggestions."""
    
    error_summary = []  # Store error messages
    
    error_summary.append("📌 **Manuav Automation Report: Issues & Actions**\n")
    error_summary.append("="*50)
    
    # ✅ 1. Missing Data Report
    missing_data_log = []
    for col in merged_df.columns:
        missing_rows = merged_df[merged_df[col].isna()]
        if not missing_rows.empty:
            for _, row in missing_rows.iterrows():
                agent = row["Agent Name"] if "Agent Name" in row else "UNKNOWN"
                date = row["Date"] if "Date" in row else "UNKNOWN DATE"
                missing_data_log.append(f"⚠ Missing data in **{col}** for **{agent}** on **{date}**.")

    if missing_data_log:
        error_summary.append("\n❌ **Missing Data Issues:**")
        error_summary.extend(missing_data_log)
        error_summary.append("💡 *Action: Please verify the source files and ensure all fields are filled.*\n")

    # ✅ 2. Time Mismatch Report
    if "Time Mismatch" in merged_df.columns:
        time_mismatch_rows = merged_df[merged_df["Time Mismatch"] != 0]
        if not time_mismatch_rows.empty:
            error_summary.append("\n⚠ **Time Mismatch Issues:**")
            for _, row in time_mismatch_rows.iterrows():
                agent = row["Agent Name"]
                date = row["Date"]
                mismatch = row["Time Mismatch"]
                error_summary.append(f"🟡 **{agent}** has a time mismatch of **{mismatch:.2f} hours** on **{date}**.")
            error_summary.append("💡 *Action: Check if agents have logged all activities correctly.*\n")

    # ✅ 3. Training Over 4 Hours Report
    if "Training" in merged_df.columns:
        high_training_rows = merged_df[merged_df["Training"] > 4]
        if not high_training_rows.empty:
            error_summary.append("\n📘 **High Training Hours Detected:**")
            for _, row in high_training_rows.iterrows():
                agent = row["Agent Name"]
                date = row["Date"]
                training_hours = row["Training"]
                error_summary.append(f"📅 **{agent}** had **{training_hours:.2f} hours** of training on **{date}**.")
            error_summary.append("💡 *Action: If this was a planned training day, no action is needed. Otherwise, verify time tracking.*\n")

    # ✅ 4. Non-Numeric Data Report
    invalid_data_log = []
    for col in ["Talk Time (Dialer) (h)", "Call Time", "Wrap-Up Time"]:
        if col in merged_df.columns:
            invalid_rows = merged_df[merged_df[col].apply(lambda x: not str(x).replace(".", "", 1).isdigit())]
            if not invalid_rows.empty:
                for _, row in invalid_rows.iterrows():
                    agent = row["Agent Name"]
                    date = row["Date"]
                    invalid_value = row[col]
                    invalid_data_log.append(f"⚠ **Invalid entry:** '{invalid_value}' in **{col}** for **{agent}** on **{date}**.")

    if invalid_data_log:
        error_summary.append("\n⚠ **Invalid Numeric Data Issues:**")
        error_summary.extend(invalid_data_log)
        error_summary.append("💡 *Action: Check if these values should be corrected or re-entered.*\n")

    # ✅ Save Error Summary to File
    error_summary_path = r"C:\Users\Emilios Richards\VSCode\Manuav Mock File Automation\reports\Error_Summary_Report.txt"
    with open(error_summary_path, "w", encoding="utf-8") as file:
        file.write("\n".join(error_summary))

    print(f"✅ Error Summary Generated and saved to: {error_summary_path}")

    return "\n".join(error_summary)  # Returning for use in emails


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

    # Standardize column names (strip spaces, fix naming issues)
    crew_df.columns = crew_df.columns.str.strip()
    dialfire_df.columns = dialfire_df.columns.str.strip()

    # Rename `User` column in Dialfire data to match Crewmeister (`Agent Name`)
    if "User" in dialfire_df.columns:
        dialfire_df.rename(columns={"User": "Agent Name"}, inplace=True)

    # Standardizing agent names for merging
    name_mapping = {
        "Anna.Jager": "Anna Jager",
        "Frederike.Tietz": "Frederike Tietz",
        "Lukas.Schneider": "Lukas Schneider",
        "Sven.Petro": "Sven Petro"
    }
    if "Agent Name" in dialfire_df.columns:
        dialfire_df["Agent Name"] = dialfire_df["Agent Name"].replace(name_mapping)

    # Convert date to standard format
    crew_df["Date"] = pd.to_datetime(crew_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    dialfire_df["Date"] = pd.to_datetime(dialfire_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

    # Merging DataFrames on Agent Name and Date
    merged_df = pd.merge(crew_df, dialfire_df, on=["Agent Name", "Date"], how="outer")

    # Ensure Training column is numeric
    merged_df["Training"] = pd.to_numeric(merged_df["Training"], errors="coerce")    

    # Ensure numeric values before calculations
    time_columns = ["Preparation Time", "Training", "Waiting Time", "Call Time", "Wrap-Up Time", "Working Time (CM)"]
    merged_df[time_columns] = merged_df[time_columns].apply(pd.to_numeric, errors="coerce")

    missing_mask_time = merged_df[time_columns].isnull()

    # Fill NaNs with 0 to prevent calculation errors
    merged_df[time_columns] = merged_df[time_columns].fillna(0)

    # Create a new column to check if Working Time (CM) matches the sum of other times
    merged_df["Time Mismatch"] = (
        merged_df["Preparation Time"] + 
        merged_df["Training"] + 
        merged_df["Waiting Time"] + 
        merged_df["Call Time"] + 
        merged_df["Wrap-Up Time"]
    ) - merged_df["Working Time (CM)"]

    # Round small floating-point errors to zero
    merged_df["Time Mismatch"] = merged_df["Time Mismatch"].apply(lambda x: 0 if abs(x) < 1e-10 else x)

    # Identify missing values before replacing them
    missing_mask = merged_df.isnull()

    # Fill missing values with 0 to prevent calculation errors
    merged_df.fillna(0, inplace=True)

    if "Talk Time (Dialer) (h)" in merged_df.columns:
        # Strip spaces and fix data type
        merged_df["Talk Time (Dialer) (h)"] = merged_df["Talk Time (Dialer) (h)"].astype(str).str.strip()

        # Convert only valid numeric values, keep non-numeric values unchanged
        merged_df["Talk Time (Dialer) (h)"] = merged_df["Talk Time (Dialer) (h)"].apply(lambda x: float(x) if x.replace(".", "", 1).isdigit() else x)

    # Function to check if a value is non-numeric
    def is_non_numeric(value):
        try:
            float(value)  # Attempt to convert to a float
            return False  # If successful, it's a valid number
        except ValueError:
            return True  # If conversion fails, it's non-numeric    

    # Save the cleaned data to Excel
    os.makedirs("reports", exist_ok=True)  # Ensure reports folder exists

        # Highlight missing data
    for column in merged_df.columns:
        if merged_df[column].isnull().any():
            print(f"Warning: Missing values found in {column}")

    
    # Save to Excel with formatting
    report_path = r"C:\Users\Emilios Richards\VSCode\Manuav Mock File Automation\reports\Merged_Automation_Report_November.xlsx"
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        merged_df.to_excel(writer, sheet_name="Merged Data", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Merged Data"]


        # ✅ 1. Apply Header Formatting (Bold, Centered, Background Color)
        format_header = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#4F81BD",  # Dark Blue Header
            "font_color": "white",
            "border": 1
        })
        

        # Apply formatting to header row
        for col_num, col_name in enumerate(merged_df.columns):
            worksheet.write(0, col_num, col_name, format_header)

        # Define the color format for training days
        format_training = workbook.add_format({"bg_color": "#A9C4EB", "font_color": "#004085"})  # Light blue background
        
        # Apply formatting to highlight missing values
        format_highlight = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        # Define the color format for invalid data (Yellow highlight)
        format_invalid = workbook.add_format({"bg_color": "#FFD966", "font_color": "#9C6500"})  # Yellow background

        for col_num, col_name in enumerate(merged_df.columns):
            worksheet.conditional_format(1, col_num, len(merged_df), col_num, {"type": "blanks", "format": format_highlight})
            missing_rows = missing_mask[col_name].values  # Find missing rows
            for row_num, is_missing in enumerate(missing_rows):
                if is_missing:
                    worksheet.write(row_num + 1, col_num, "MISSING", format_highlight)  # Highlight missing cell

        # Apply formatting to highlight entire row if Training > 4 hours
        for row_num in range(1, len(merged_df) + 1):  # Iterate over rows
            if merged_df.loc[row_num - 1, "Training"] > 4.0:  # If training exceeds 4 hours
                for col_num in range(len(merged_df.columns)):  # Iterate over all columns
                    worksheet.write(row_num, col_num, merged_df.iloc[row_num - 1, col_num], format_training)  # Apply blue format
        
        # Apply formatting to highlight time mismatches
        if "Time Mismatch" in merged_df.columns:    
            for row_num in range(1, len(merged_df) + 1):  # Iterate over rows
                if merged_df.loc[row_num - 1, "Time Mismatch"] != 0:  # If mismatch exists
                    worksheet.write(row_num, merged_df.columns.get_loc("Time Mismatch"), merged_df.loc[row_num - 1, "Time Mismatch"], format_highlight)
        else:
            print("❌ ERROR: 'Time Mismatch' column is missing from merged_df!")

        # Apply formatting to highlight invalid numeric values in time-based columns
        time_columns_with_talk = ["Preparation Time", "Training", "Waiting Time", "Call Time", "Wrap-Up Time", "Working Time (CM)", "Talk Time (Dialer) (h)"]

        # Identify rows where "Talk Time (Dialer) (h)" is non-numeric BEFORE converting
        if "Talk Time (Dialer) (h)" in merged_df.columns:
            non_numeric_mask = merged_df["Talk Time (Dialer) (h)"].apply(lambda x: pd.notna(x) and not str(x).replace(".", "", 1).isdigit())  # Detect non-numeric values
            talk_time_numeric = pd.to_numeric(merged_df["Talk Time (Dialer) (h)"], errors="coerce")  # Convert numeric values to a temporary variable

        # Define a yellow format for invalid data
        format_invalid = workbook.add_format({"bg_color": "#FFD966", "font_color": "#9C6500"})  # Yellow background

        # Highlight non-numeric "Talk Time (Dialer) (h)" values
        if "Talk Time (Dialer) (h)" in merged_df.columns:
            for row_num in range(1, len(merged_df) + 1):  # Iterate over rows
                value = merged_df.loc[row_num - 1, "Talk Time (Dialer) (h)"]
                if non_numeric_mask.iloc[row_num - 1]:  # If value was non-numeric before conversion
                    worksheet.write(row_num, merged_df.columns.get_loc("Talk Time (Dialer) (h)"), merged_df.loc[row_num - 1, "Talk Time (Dialer) (h)"], format_invalid)  # Highlight in yellow

        for col_name in time_columns_with_talk:
            if col_name in merged_df.columns:
                for row_num in range(1, len(merged_df) + 1):  # Iterate over rows
                    value = merged_df.loc[row_num - 1, col_name]
                    if is_non_numeric(value):  # If value is non-numeric
                        worksheet.write(row_num, merged_df.columns.get_loc(col_name), value, format_invalid)  # Highlight in yellow
 
        # Highlight missing values in time-based columns
        for col_num, col_name in enumerate(time_columns):
            for row_num in range(1, len(merged_df) + 1):  # Iterate over rows
                if missing_mask.loc[row_num - 1, col_name]:  # If originally missing
                    worksheet.write(row_num, col_num, merged_df.loc[row_num - 1, col_name], format_invalid)

        # Define a bold format for column headers
        format_bold = workbook.add_format({"bold": True})

        # Add a bold header label to the "Time Mismatch" column
        worksheet.write(0, merged_df.columns.get_loc("Time Mismatch"), "Time Mismatch (Error Check)", format_bold)

        # After writing merged_df to the Excel file:
        for i, col in enumerate(merged_df.columns):
            # Determine the maximum length in the column (including header)
            max_length = max(
                merged_df[col].astype(str).map(len).max(), 
                len(col)
            )
            # Set the column width, adding a little extra space
            worksheet.set_column(i, i, max_length + 2)
            
    
    print(f"✅ Automation Report Generated Successfully! Saved to {report_path}")
    
    # ✅ Generate and Save Error Summary Report
    generate_error_summary(merged_df)



if __name__ == "__main__":
    merge_and_clean_data()