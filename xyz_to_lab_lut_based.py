#XYZ TO LAB LUT BASED
import pandas as pd
import numpy as np
from scipy.interpolate import griddata

# Read the LUT data from the first Excel file
lut_file_path = 'lut_lab_to_xyz.xlsx'
lut_data = pd.read_excel(lut_file_path)

# Extract the columns for interpolation
XYZ_values_lut = lut_data[['X', 'Y', 'Z']]
lab_values_lut = lut_data[['L', 'a', 'b']]


# Read the XYZ values from the second Excel file
input_file_path = 'rgb2xyz_data.xlsx'  # Replace with the path to your second Excel file
input_data = pd.read_excel(input_file_path)

# Extract the XYZ values as a 2D NumPy array
new_XYZ_values = input_data[['X', 'Y', 'Z']].values
pixel_numbers = input_data['PixelNumber'].values  # Extract the Pixel Number column

# Perform interpolation using NumPy vectorized operations
interpolated_lab_values = griddata(XYZ_values_lut, lab_values_lut, new_XYZ_values, method='linear')

# Create a DataFrame for the interpolated L*a*b* values
output_df = pd.DataFrame(interpolated_lab_values, columns=['L*', 'a*', 'b*'])

# Insert the 'Pixel Number' column as the first column in the DataFrame
output_df.insert(0, 'PixelNumber', pixel_numbers)

# Save the DataFrame to an Excel file
output_file_path = 'output_for_luts_XYZ_to_Lab.xlsx'  # Replace with the desired path for the output Excel file
output_df.to_excel(output_file_path, index=False)

