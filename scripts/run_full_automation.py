import subprocess

print("🚀 Running Full Automation Pipeline...\n")

# ✅ Step 1: Merge & Clean Data
print("🔄 Merging and cleaning data...")
subprocess.run(["python", "scripts/merge_reports.py"])

# ✅ Step 2: Generate Visualizations
print("\n📊 Generating data visualizations...")
subprocess.run(["python", "scripts/generate_visuals.py"])

# ✅ Step 3: Send Email Report
print("\n📩 Sending automated email report...")
subprocess.run(["python", "scripts/send_email.py"])

print("\n✅ Full Automation Completed Successfully!")
