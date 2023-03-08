# Import Libraries 
import pandas as pd 
import os

# Set up name of new printed Excel report
project_name = input("Project Name: ")
axis_name = input("Axis Name: ")

# Pandas Read
location = r"I:\MECH\SHA\Analyzer Selection.xlsx"
df = pd.read_excel(location)

# Read Gearbox Specs from Excel Database 
gbx_item_num = pd.to_numeric(df.iloc[1:,0], errors='coerce').apply(lambda x: int(x))
gbx_name = df.iloc[1:, 1].astype(str).str.strip()
gbx_frame = pd.to_numeric(df.iloc[1:,2], errors='coerce')
gbx_ratio = pd.to_numeric(df.iloc[1:,3], errors='coerce')
gbx_speed = pd.to_numeric(df.iloc[1:,4], errors='coerce')
gbx_torque = pd.to_numeric(df.iloc[1:,5], errors='coerce')
gbx_cost = pd.to_numeric(df.iloc[1:,6], errors='coerce')

# Inputs
gbx_type = input("\nSelect type of gearbox: \n NPL \n Cone Drive \n NVH \n VDH \n VH+ \n\n SELECT: ")

if gbx_type == "Cone Drive":
    first_three_letters = "S"
else:
    first_three_letters = gbx_type[:4]  # Get the first two letters of the selected gearbox type

# Filter the gbx_name column based on the first two letters of gearbox names
selected_gbx_names = gbx_name[gbx_name.str.startswith(first_three_letters)]

# Error Handling 
if len(selected_gbx_names) == 0:
    print("No gearbox names found for the selected type.")
else:
    print(selected_gbx_names)
    
peak_speed = float(input("\nPeak Speed: "))


# Calculate the new speed for each selected gearbox
new_speeds = []
for ratio in gbx_ratio[gbx_name.str.startswith(first_three_letters)]:
    new_speeds.append(ratio * peak_speed)

# Read Servo Specs from Excel Database 
servo_item_num = pd.to_numeric(df.iloc[1:,9], errors='coerce').dropna().apply(lambda x: int(x))
servo_name = df.iloc[1:, 10].astype(str).str.strip().dropna()
servo_cont_torque = pd.to_numeric(df.iloc[1:,11], errors='coerce').dropna()
servo_peak_torque = pd.to_numeric(df.iloc[1:,12], errors='coerce').dropna()
servo_velocity = pd.to_numeric(df.iloc[1:,13], errors='coerce').dropna()
servo_cost = pd.to_numeric(df.iloc[1:,15], errors='coerce').dropna()

# Create a list of dictionaries with the desired columns
results = []
for name, ratio in zip(selected_gbx_names, gbx_ratio[gbx_name.str.startswith(first_three_letters)]):
    new_speed = ratio * peak_speed
    for servo, velocity, cost in zip(servo_name, servo_velocity, servo_cost):
        percentage = new_speed / velocity * 100
        if percentage <= 80:
            total_cost = gbx_cost[gbx_name == name].iloc[0] + cost
            result = {"Servo": servo, "Gearbox Name": name, "Percentage": percentage, "Total Cost": total_cost}
            results.append(result)

# Convert the list of dictionaries to a pandas dataframe
df_results = pd.DataFrame(results)

# Export the dataframe to Excel
file_name = f"{project_name}.xlsx"
sheet_name = axis_name
df_results.to_excel(file_name, sheet_name=sheet_name, index=False)

# Set the location of the export
export_location = r"C:\Users\shacosta\Desktop"
file_path = os.path.join(export_location, file_name)
df_results.to_excel(file_path, sheet_name=sheet_name, index=False)
