import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the merged dataset
file_path = r"C:\Users\Emilios Richards\VSCode\Manuav Mock File Automation\reports\Merged_Automation_Report_November.xlsx"
merged_df = pd.read_excel(file_path)


# ✅ 1. Bar Chart: Total Training Hours Per Agent
plt.figure(figsize=(10, 5))
# Drop rows where Training is missing
training_df = merged_df.dropna(subset=["Agent Name", "Training"])

sns.barplot(data=training_df, x="Agent Name", y="Training", errorbar=None, hue="Agent Name", palette="Blues_d", legend=False)
plt.xticks(rotation=45)
plt.title("Total Training Hours per Agent")
plt.xlabel("Agent Name")
plt.ylabel("Training Hours")
plt.tight_layout()
plt.savefig(r"C:\Users\Emilios Richards\VSCode\Manuav Mock File Automation\visuals\training_hours.png")  # Save to file
plt.show()


# ✅ 2. Line Chart: Time Mismatch Trends Over Time
# Drop rows where Time Mismatch is missing
if "Time Mismatch (Error Check)" in merged_df.columns:
    mismatch_df = merged_df.dropna(subset=["Date", "Time Mismatch (Error Check)"])
    
    plt.figure(figsize=(10, 5))
    sns.lineplot(data=mismatch_df, x="Date", y="Time Mismatch (Error Check)", marker="o", color="red")
    plt.xticks(rotation=45)
    plt.title("Time Mismatch Trends Over Time")
    plt.xlabel("Date")
    plt.ylabel("Time Mismatch (Hours)")
    plt.grid()
    plt.tight_layout()
    plt.savefig(r"C:\Users\Emilios Richards\VSCode\Manuav Mock File Automation\visuals\time_mismatch_trends.png")
    plt.show()
else:
    print("❌ ERROR: 'Time Mismatch' column not found in DataFrame!")

# ✅ 3. Pie Chart: Call Success vs. Failures
if "Successful" in merged_df.columns and "Completed" in merged_df.columns:
    # Fill missing values with 0 before summing
    call_data = merged_df[["Successful", "Completed"]].fillna(0).sum()

    plt.figure(figsize=(5, 5))
    plt.pie(call_data, labels=["Successful Calls", "Failed Calls"], autopct="%1.1f%%", colors=["green", "red"])
    plt.title("Call Success vs. Failure Rate")
    plt.savefig(r"C:\Users\Emilios Richards\VSCode\Manuav Mock File Automation\visuals\call_performance.png")
    plt.show()
else:
    print("❌ ERROR: 'Successful' or 'Completed' column not found in DataFrame!")

print("✅ Visualizations Generated & Saved to Reports Folder!")
