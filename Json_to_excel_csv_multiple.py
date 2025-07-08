
import json
import pandas as pd
import os

# Define your input folder
root_folder = 'Jsonfiles_input'
# Define your output folder
output_folder = r"Jsonfiles_output"

# Ensure the folder exists (create if not)
os.makedirs(output_folder, exist_ok=True)

# Combine path with filename
output_excel_file = os.path.join(output_folder, 'multi_sheet_nested_output.xlsx')

# Create Excel writer
with pd.ExcelWriter(output_excel_file, engine='openpyxl') as writer:
    for dirpath, dirnames, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.endswith('.json'):
                file_path = os.path.join(dirpath, filename)
                try:
                    with open(file_path, 'r') as f:
                        data = json.load(f)
                        if isinstance(data, list):
                            df = pd.DataFrame(data)

                            # ---- Excel Output ----
                            relative_path = os.path.relpath(file_path, root_folder)
                            sheet_name = os.path.splitext(relative_path.replace("\\", "_").replace("/", "_"))[0][:31]
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                              # ---- CSV Output ----
                            csv_filename = os.path.splitext(filename)[0] + '.csv'
                            csv_output_path = os.path.join(output_folder, csv_filename)
                            df.to_csv(csv_output_path, index=False)
                        else:
                            print(f"Skipped {file_path}: not a list of records.")
                except Exception as e:
                    print(f"Error reading {file_path}: {e}")

print(f"✅ Exported Excel to: {output_excel_file}")
print(f"✅ CSV files saved in: {output_folder}")
