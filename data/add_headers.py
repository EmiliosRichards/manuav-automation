import pandas as pd

# Define the file path (relative to your main project folder)
file_path = "crewmeister.xlsx"

# Load the Excel file, assuming the actual data starts after 1 or 2 rows
df = pd.read_excel(file_path, skiprows=1)  # Adjust `skiprows` if necessary

# Transpose the DataFrame so that rows become columns
df = df.T  # Equivalent to df.transpose()

# Assign the first row as new column headers
df.columns = df.iloc[0]  # Set first row as column headers
df = df[1:]  # Remove the old header row

# Reset index to clean up
df.reset_index(drop=True, inplace=True)

# Identify the columns that should contain "Alias" and "Agent Name"
df.rename(columns={df.columns[0]: "Alias", df.columns[1]: "Agent Name"}, inplace=True)

# Forward fill missing alias and agent names (to ensure they match the correct rows)
df["Alias"].fillna(method="ffill", inplace=True)
df["Agent Name"].fillna(method="ffill", inplace=True)


# Save to a new Excel file
output_file = "restructured_crewmeister.xlsx"
df.to_excel(output_file, index=False)


print(f"✅ Header added successfully! File saved as {output_file}")

print(df.head())  # Check the first few rows