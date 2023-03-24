# Import Libraries 
import pandas as pd 
import os

# Set up name of new printed Excel report
project_name = "BRENTON 4482"
axis_name = "LOADER"

# Pandas Read
location = r"C:\Users\shacosta\Desktop\Analyzer Selection2.xlsx"
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
gbx_type = "NPL"

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


# Input Calculations from Motion Analyzer 
peak_speed = 280.74
peak_torque_input = 66.65
peak_acceleration = 102.83
max_inertia = 0.49


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

seen_combinations = set()
# Create a list of dictionaries with the desired columns
results = []
for name, ratio in zip(selected_gbx_names, gbx_ratio[gbx_name.str.startswith(first_three_letters)]):
    new_speed = ratio * peak_speed
    for servo, velocity, cost, peak_torque, cont_torque in zip(servo_name, servo_velocity, servo_cost, servo_peak_torque, servo_cont_torque):
        # Calculations 
        post_gbx_torque = (peak_torque_input) / ((ratio) * ((1 - ((max_inertia* peak_acceleration) / (peak_torque * ratio)))))
        percentage_speed = (new_speed / velocity) * 100
        percentage_peak_torque = (post_gbx_torque / peak_torque) * 100
        percentage_cont_torque = (post_gbx_torque / cont_torque) * 100
        
        gbx_speed_protect = ( (peak_speed  * ratio) / gbx_speed) * 100
        gbx_torque_protect = (peak_torque_input / gbx_torque) * 100

        combination = None
        
        mask = (percentage_speed <= 80) & (percentage_peak_torque <= 80) & (percentage_cont_torque <= 65) & (gbx_speed_protect <= 80) & (gbx_torque_protect <= 80) & (percentage_peak_torque >= 0) & (percentage_cont_torque >= 0)
        if mask.any():
            combination = (percentage_speed, percentage_peak_torque, percentage_cont_torque)        
            
            if combination is not None and combination not in seen_combinations:
                total_cost = cost + gbx_cost[gbx_name == name].iloc[0]
                seen_combinations.add(combination)
                result = {"Servo": servo, "Gearbox Name": name, "Percentage Speed": percentage_speed, "Percentage Peak Torque": percentage_peak_torque, "Percentage Cont Torque": percentage_cont_torque, "Total Cost": total_cost}
                results.append(result)

# Convert the list of dictionaries to a pandas dataframe
df_results = pd.DataFrame(results)

print(df_results)
