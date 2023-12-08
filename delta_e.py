#Delta E calculation
import pandas as pd
import numpy as np

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

for i in range(len(merged_lab_data)):
    lab1 = merged_lab_data.loc[i, ['L*_1', 'a*_1', 'b*_1']].values
    lab2 = merged_lab_data.loc[i, ['L*_2', 'a*_2', 'b*_2']].values

    delta_e_76 = np.sqrt(np.sum((lab1 - lab2) ** 2))

    delta_e_76_values.append(delta_e_76)

# Add the calculated delta E values to the DataFrame
merged_lab_data['DeltaE_76'] = delta_e_76_values

# Reorder the columns
merged_lab_data = merged_lab_data[['PixelNumber', 'L*_1', 'a*_1', 'b*_1', 'L*_2', 'a*_2', 'b*_2', 'DeltaE_76']]

# Create a Pandas Excel writer using XlsxWriter as the engine
excel_writer = pd.ExcelWriter('delta_e_values_RGB_to_Lab_LUT.xlsx', engine='xlsxwriter')

# Write the DataFrame to a sheet in the Excel file
merged_lab_data.to_excel(excel_writer, index=False)

# Save the Excel file
excel_writer.save()

avg = merged_lab_data.loc[:,'DeltaE_76'].mean()
print('The average Delta E 76 is:', avg)
