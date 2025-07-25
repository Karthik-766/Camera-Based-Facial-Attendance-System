import pandas as pd

file_path = "attendance.xlsx"

# Load the attendance data
df = pd.read_excel(file_path)

# Set the date threshold
date_threshold = "2025-03-29"

# Filter records (keep only recent entries)
df = df[df["Date"] >= date_threshold]

# Save the updated file
df.to_excel(file_path, index=False)

print(f"Deleted records before {date_threshold}!")
