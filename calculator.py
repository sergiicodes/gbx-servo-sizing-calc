# Import Libraries 
import pandas as pd 
import numpy as np
#import os

'''
# Set up name of new printed Excel report
project_name = "BRENTON 4482"
axis_name = "LOADER"
'''

# Pandas Read
location = r"C:\Users\jbooker\Desktop\Analyzer Selection.xlsx"
df_gbx = pd.read_excel(location, sheet_name="GBX DATABASE")
df_servo = pd.read_excel(location, sheet_name="SERVO DATABASE")

# Read Gearbox Specs from Excel Database 
gbx_item_num = pd.to_numeric(df_gbx.iloc[1:,0], errors='coerce').apply(lambda x: int(x))
gbx_name = df_gbx.iloc[1:, 1].astype(str).str.strip()
gbx_ratio = pd.to_numeric(df_gbx.iloc[1:,3], errors='coerce')
gbx_speed = pd.to_numeric(df_gbx.iloc[1:,4], errors='coerce')
gbx_torque = pd.to_numeric(df_gbx.iloc[1:,5], errors='coerce')
gbx_cost = pd.to_numeric(df_gbx.iloc[1:,6], errors='coerce')

# Inputs
#gbx_type = input("\nSelect type of gearbox: \n NPL \n Cone Drive \n NVH \n VDH \n VH+ \n TP \n\n SELECT: ")
gbx_type = "NPL"

if gbx_type == "Cone Drive":
    first_three_letters = "S"
else:
    first_three_letters = gbx_type[:4]  # Get the first two letters of the selected gearbox type

# Filter the gbx_name column based on the first two letters of gearbox names
selected_gbx_names = gbx_name[gbx_name.str.startswith(first_three_letters)]

print(selected_gbx_names)

# Input Calculations from Motion Analyzer 
peak_speed = 117.76
peak_torque_input = 5.44
peak_acceleration = 471.06
max_inertia = 707.91 / 10000

# Calculate the new speed for each selected gearbox
new_speeds = []
for ratio in gbx_ratio[gbx_name.str.startswith(first_three_letters)]:
    new_speeds.append(ratio * peak_speed)


# Read Servo Specs from Excel Database 
servo_item_num = pd.to_numeric(df_servo.iloc[1:,0], errors='coerce').dropna().apply(lambda x: int(x))
servo_name = df_servo.iloc[1:, 1].astype(str).str.strip().dropna()
servo_cont_torque = pd.to_numeric(df_servo.iloc[1:,2], errors='coerce').dropna()
servo_peak_torque = pd.to_numeric(df_servo.iloc[1:,3], errors='coerce').dropna()
servo_velocity = pd.to_numeric(df_servo.iloc[1:,4], errors='coerce').dropna()
servo_cost = pd.to_numeric(df_servo.iloc[1:,6], errors='coerce').dropna()


seen_combinations = set()
# Create a list of dictionaries with the desired columns
results = []
for name, ratio, gbxSpeed, gbxTq in zip(selected_gbx_names, gbx_ratio[gbx_name.str.startswith(first_three_letters)], gbx_speed[gbx_name.str.startswith(first_three_letters)], gbx_torque[gbx_name.str.startswith(first_three_letters)]):
    new_speed = ratio * peak_speed
    for servo, velocity, cost, peak_torque, cont_torque in zip(servo_name, servo_velocity, servo_cost, servo_peak_torque, servo_cont_torque):
        # Calculations 
        post_gbx_torque = (peak_torque_input) / ((ratio) * ((1 - ((max_inertia* peak_acceleration) / (peak_torque * ratio)))))
        POST_GBX_TORQUE = ( (peak_acceleration*(np.pi/30)) * max_inertia ) 
        
        servo_percentage_speed = (new_speed / velocity) * 100
        servo_percentage_peak_torque = (POST_GBX_TORQUE / peak_torque) * 100
        #percentage_cont_torque = (post_gbx_torque / cont_torque) * 100
        
        gbx_speed_protect = ( (peak_speed  * ratio) / gbxSpeed) * 100
        gbx_torque_protect = ( peak_torque_input / gbxTq) * 100
        
        combination = None
        
        #mask = (percentage_speed <= 80) & (percentage_peak_torque <= 80) & (percentage_cont_torque <= 70) & (gbx_speed_protect <= 80) & (gbx_torque_protect <= 80) & (percentage_peak_torque >= 0) & (percentage_cont_torque >= 0)
        mask = (servo_percentage_speed <= 80) & (servo_percentage_peak_torque <= 80) & (gbx_speed_protect <= 80) & (gbx_torque_protect <= 80) & (servo_percentage_peak_torque >= 0) 
        if (servo_percentage_speed <= 80) and (servo_percentage_peak_torque <= 80) and (gbx_speed_protect <= 80) and (gbx_torque_protect <= 80) and (servo_percentage_peak_torque >= 0):
            combination = (servo_percentage_speed, servo_percentage_peak_torque)      

            if combination is not None and combination not in seen_combinations:
                total_cost = cost + gbx_cost[gbx_name == name].iloc[0]
                seen_combinations.add(combination)
                #result = {"Servo": servo, "Gearbox Name": name, "Percentage Speed [Servo]": percentage_speed, "Percentage Peak Torque [Servo]": percentage_peak_torque, "Percentage Cont Torque [Servo]": percentage_cont_torque, "Percentage Speed [GBX]": gbx_speed_protect, "Percentage Torque [GBX]": gbx_torque_protect, "Total Cost": total_cost}
                result = {"Servo": servo, "Gearbox Name": name, "Percentage Speed [Servo]": servo_percentage_speed, "Percentage Peak Torque [Servo]": servo_percentage_peak_torque, "Percentage Speed [GBX]": gbx_speed_protect, "Percentage Torque [GBX]": gbx_torque_protect, "Total Cost": total_cost}
                results.append(result)
                
# Convert the list of dictionaries to a pandas dataframe
df_results = pd.DataFrame(results)

print(df_results)
