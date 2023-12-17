#Delta E ALL
import pandas as pd
import numpy as np

def calculate_delta_e_94(lab1, lab2, k_L=1, k_C=1, k_H=1):
    delta_L = lab2[0] - lab1[0]
    delta_a = lab2[1] - lab1[1]
    delta_b = lab2[2] - lab1[2]
    C_1 = np.sqrt(lab1[1] ** 2 + lab1[2] ** 2)
    C_2 = np.sqrt(lab2[1] ** 2 + lab2[2] ** 2)
    delta_C_ab = C_2 - C_1
    delta_H_ab = np.sqrt(delta_a**2 + delta_b**2 - delta_C_ab**2)
    s_L = 1
    s_C = 1 + k_C * C_1
    s_H = 1 + k_H * C_1
    delta_E_94 = np.sqrt((delta_L / (k_L * s_L)) ** 2 + (delta_C_ab / (k_C * s_C)) ** 2 + (delta_H_ab / (k_H * s_H)) ** 2)
    return delta_E_94

def calculate_delta_e_2000(lab1, lab2, k_L=1, k_C=1, k_H=1):
    L_1, a_1, b_1 = lab1
    L_2, a_2, b_2 = lab2

    delta_L = L_2 - L_1
    L_bar = (L_1 + L_2) / 2
    C_1 = np.sqrt(a_1**2 + b_1**2)
    C_2 = np.sqrt(a_2**2 + b_2**2)
    C_bar = (C_1 + C_2) / 2
    G = 0.5 * (1 - np.sqrt(C_bar**7 / (C_bar**7 + 25**7)))

    a_1_prime = (1 + G) * a_1
    a_2_prime = (1 + G) * a_2

    C_1_prime = np.sqrt(a_1_prime**2 + b_1**2)
    C_2_prime = np.sqrt(a_2_prime**2 + b_2**2)
    C_bar_prime = (C_1_prime + C_2_prime) / 2
    delta_C_prime = C_2_prime - C_1_prime

    h_1_prime = np.arctan2(b_1, a_1_prime)
    h_2_prime = np.arctan2(b_2, a_2_prime)

    delta_h_prime = np.where(np.abs(h_1_prime - h_2_prime) <= np.pi, h_2_prime - h_1_prime, h_2_prime - h_1_prime + 2*np.pi)

    delta_H_prime = 2 * np.sqrt(C_1_prime * C_2_prime) * np.sin(delta_h_prime / 2)

    S_L = 1 + (k_L * (L_bar - 50)) / 100
    S_C = 1 + (k_C * C_bar_prime) / 100
    S_H = 1 + (k_H * C_bar_prime) / 100

    delta_theta = np.radians(30 * np.exp(-((180 / np.pi * np.arccos(np.cos(np.radians((h_1_prime + h_2_prime) / 2))) - 275) / 25)**2))

    R_C = 2 * np.sqrt(C_bar_prime**7 / (C_bar_prime**7 + 25**7))

    R_T = -np.sin(2 * delta_theta) * R_C

    delta_E_2000 = np.sqrt((delta_L / (k_L * S_L))**2 + (delta_C_prime / (k_C * S_C))**2 + (delta_H_prime / (k_H * S_H))**2 + R_T * (delta_C_prime / (k_C * S_C)) * (delta_H_prime / (k_H * S_H)))

    return delta_E_2000

# Read the first set of L*a*b* values from the first Excel file
lab_file1_path = 'output_for_luts.xlsx'
lab_data1 = pd.read_excel(lab_file1_path)

# Read the second set of L*a*b* values from the second Excel file
lab_file2_path = 'output_for_luts_XYZ_to_Lab.xlsx'
lab_data2 = pd.read_excel(lab_file2_path)

# Merge the two DataFrames on the pixel number
merged_lab_data = lab_data1.merge(lab_data2, on='PixelNumber', suffixes=('_1', '_2'))

# Calculate the delta E values
delta_e_76_values = []
delta_e_94_values = []
delta_e_2000_values = []

for i in range(len(merged_lab_data)):
    lab1 = merged_lab_data.loc[i, ['L*_1', 'a*_1', 'b*_1']].values
    lab2 = merged_lab_data.loc[i, ['L*_2', 'a*_2', 'b*_2']].values

    delta_e_76 = np.sqrt(np.sum((lab1 - lab2) ** 2))
    delta_e_94 = calculate_delta_e_94(lab1, lab2)
    delta_e_2000 = calculate_delta_e_2000(lab1, lab2)

    delta_e_76_values.append(delta_e_76)
    delta_e_94_values.append(delta_e_94)
    delta_e_2000_values.append(delta_e_2000)

# Add the calculated delta E values to the DataFrame
merged_lab_data['DeltaE_76'] = delta_e_76_values
merged_lab_data['DeltaE_94'] = delta_e_94_values
merged_lab_data['DeltaE_2000'] = delta_e_2000_values

# Reorder the columns
merged_lab_data = merged_lab_data[['PixelNumber', 'L*_1', 'a*_1', 'b*_1', 'L*_2', 'a*_2', 'b*_2', 'DeltaE_76', 'DeltaE_94', 'DeltaE_2000']]

# Create a Pandas Excel writer using XlsxWriter as the engine
excel_writer = pd.ExcelWriter('delta_e_values_RGB_to_Lab_LUT.xlsx', engine='xlsxwriter')

# Write the DataFrame to a sheet in the Excel file
merged_lab_data.to_excel(excel_writer, index=False)

# Save the Excel file
excel_writer.save()

avg_76 = merged_lab_data.loc[:, 'DeltaE_76'].mean()
avg_94 = merged_lab_data.loc[:, 'DeltaE_94'].mean()
avg_2000 = merged_lab_data.loc[:, 'DeltaE_2000'].mean()

print('The average Delta E 76 is:', avg_76)
print('The average Delta E 94 is:', avg_94)
print('The average Delta E 2000 is:', avg_2000)
