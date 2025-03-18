import pandas as pd
import random

# Simulated Crewmeister Data (Work Hours)
crewmeister_data = {
    "Agent Name": ["John Doe", "Jane Smith", "Alice Brown", "Bob White"],
    "Date": ["2024-03-14"] * 4,
    "Total Hours": [7.5, 6.2, 8.0, 7.0],
    "Clock-in Time": ["08:00", "09:15", "07:45", "08:30"],
    "Clock-out Time": ["15:30", "15:25", "16:00", "15:45"]
}
crewmeister_df = pd.DataFrame(crewmeister_data)

# Simulated Dialfire Call Statistics (Call Logs)
dialfire_data = {
    "Agent Name": ["J. Doe", "Jane S.", "A. Brown", "Bob W."],  # Variance in agent names
    "Date": ["2024-03-14"] * 4,
    "Calls Made": [50, 40, 60, 55],
    "Successful Calls": [10, 7, 12, 9],
    "Call Duration (mins)": [120, 100, 150, 140]
}
dialfire_df = pd.DataFrame(dialfire_data)

# Standardizing Agent Names (Fixing Variance)
name_mapping = {
    "J. Doe": "John Doe",
    "Jane S.": "Jane Smith",
    "A. Brown": "Alice Brown",
    "Bob W.": "Bob White"
}
dialfire_df["Agent Name"] = dialfire_df["Agent Name"].replace(name_mapping)

# Merging Data on Agent Name and Date
merged_df = pd.merge(crewmeister_df, dialfire_df, on=["Agent Name", "Date"], how="outer")

# Save to Excel
merged_df.to_excel("Manuav_Automation_Report.xlsx", index=False)
print("Automation Report Generated Successfully!")
