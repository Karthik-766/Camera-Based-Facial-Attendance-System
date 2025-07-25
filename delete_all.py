import pandas as pd

file_path = "attendance.xlsx"

# Create an empty DataFrame with headers
df = pd.DataFrame(columns=["Name", "Time", "Date"])

# Save the empty DataFrame to Excel (clears all data)
df.to_excel(file_path, index=False)

print("All attendance records deleted!")
