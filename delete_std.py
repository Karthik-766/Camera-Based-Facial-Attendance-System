import pandas as pd

file_path = "attendance.xlsx"
student_to_remove = "Pravallika"

df = pd.read_excel(file_path)
df = df[df["Name"] != student_to_remove]  # Keep all except the specified student

df.to_excel(file_path, index=False)
print(f"Deleted attendance records for {student_to_remove}!")
