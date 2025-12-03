
import pandas as pd
import subprocess
import time

# ==============================
# CONFIGURATION
# ==============================
excel_file = r"C:/Users/JOSHINMA/Documents/MyCode/AutomationCM/IntegrityItemID.xlsx"
output_file = f"C:/Users/JOSHINMA/Documents/MyCode/AutomationCM/integrity_data_filled_{int(time.time())}.xlsx"

# ==============================
# READ EXCEL
# ==============================
df = pd.read_excel(excel_file)

if 'ItemID' not in df.columns:
    raise ValueError("Excel must have 'ItemID' as the first column.")

# Convert all non-ID columns to string type
for col in df.columns:
    if col != 'ItemID':
        df[col] = df[col].astype(str)

# ==============================
# PROCESS EACH ITEM
# ==============================
for index, row in df.iterrows():
    try:
        item_id = str(int(row['ItemID']))
        attributes = [col for col in df.columns if col != 'ItemID']
        
        # Build CLI command with delimiter
        fields = ",".join(attributes)
        cmd = ["im", "issues", f"--fieldsDelim=|", f"--fields={fields}", item_id]
        
        # Run CLI command
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            output = result.stdout.strip()
            print(f"\nRaw CLI Output for {item_id}:\n{output}")
            
            lines = [l for l in output.splitlines() if l.strip()]
            
            if len(lines) == 1:
                # Split by delimiter
                values = lines[0].split("|")
                if len(values) == len(attributes):
                    data_map = dict(zip(attributes, values))
                    print(f"Parsed Data for {item_id}: {data_map}")
                    for attr in attributes:
                        df.at[index, attr] = data_map[attr]
                else:
                    print(f"⚠ Unexpected number of values for Item ID {item_id}")
            else:
                print(f"⚠ No data returned for Item ID {item_id}")
        else:
            print(f" Error fetching data for Item ID {item_id}: {result.stderr.strip()}")
    
    except Exception as e:
        print(f"⚠ Exception for Item ID {item_id}: {e}")

# ==============================
# SAVE UPDATED EXCEL
# ==============================
df.to_excel(output_file, index=False)
print(f"\n✅ Data fetching completed! Saved as {output_file}")
