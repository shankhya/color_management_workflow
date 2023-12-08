#LAB TO XYZ mathematical
import pandas as pd
import numpy as np

# Load LAB data from Excel
lab_file_path = 'output_for_luts.xlsx'  # Replace with your LAB data file path
lab_data = pd.read_excel(lab_file_path)

# Extract LAB values
L_values = lab_data['L*']
a_values = lab_data['a*']
b_values = lab_data['b*']

# Reference white point for D65 illuminant
Xn, Yn, Zn = 0.95047, 1.00000, 1.08883


# Define the conversion functions
def lab_to_xyz(L, a, b):
    delta = 6 / 29
    fy = (L + 16) / 116
    fx = (a / 500) + fy
    fz = fy - (b / 200)

    def compand(value):
        return value ** 3 if value > delta else (value - 16 / 116) / 7.787

    X = Xn * compand(fx)
    Y = Yn * compand(fy)
    Z = Zn * compand(fz)

    return X, Y, Z


# Apply conversion to all rows in the DataFrame
xyz_data = lab_data.apply(lambda row: pd.Series(lab_to_xyz(row['L*'], row['a*'], row['b*'])), axis=1)
xyz_data.columns = ['X', 'Y', 'Z']

# Concatenate the original LAB data with the new XYZ data
result_data = pd.concat([lab_data['PixelNumber'], xyz_data], axis=1)

# Save the result to a new Excel file
output_file_path = 'output_lab_to_xyz_manual.xlsx'  # Replace with your desired output file path
result_data.to_excel(output_file_path, index=False)
