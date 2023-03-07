import pandas as pd 

# pandas read
location = r"I:\MECH\SHA\Analyzer Selection.xlsx"
df = pd.read_excel(location)

# Gearbox Section
gbx_item_num = pd.to_numeric(df.iloc[1:,0], errors='coerce')
gbx_name = df.iloc[1:, 1].astype(str).str.strip()
gbx_frame = pd.to_numeric(df.iloc[1:,2], errors='coerce')
gbx_ratio = pd.to_numeric(df.iloc[1:,3], errors='coerce')
gbx_speed = pd.to_numeric(df.iloc[1:,4], errors='coerce')
gbx_torque = pd.to_numeric(df.iloc[1:,5], errors='coerce')
gbx_cost = pd.to_numeric(df.iloc[1:,6], errors='coerce')


# Inputs
gbx_type = input("Select type of gearbox: \n NPL \n Cone Drive \n NVH \n VDH \n VH+ \n\n SELECT: ")

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
'''
peak_torque = float(input("\nPeak Torque [Nm]: "))

peak_accel = float(input("\nPeak Acceleration: "))
'''

for name, ratio in zip(selected_gbx_names, gbx_ratio[gbx_name.str.startswith(first_three_letters)]):
    new_speed = ratio * peak_speed
    print(f"{name}: {new_speed:.2f} RPM")

'''
possible_ratios = peak_accel * ratios
print(possible_ratios)
'''
