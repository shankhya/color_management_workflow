import pandas as pd
import numpy as np
def xyz_to_lab(xyz):
    # Reference white point for D65 illuminant
    Xn, Yn, Zn = 0.950456, 1.0, 1.088754
    # Normalize XYZ values
    xn = xyz[0] / Xn
    yn = xyz[1] / Yn
    zn = xyz[2] / Zn

    # Non-linear functions
    def f(t):
        return t**(1/3) if t > 0.008856 else 903.3 * t + 16/116

    # Convert XYZ to L*a*b*
    L = 116 * f(yn) - 16
    a = 500 * (f(xn) - f(yn))
    b = 200 * (f(yn) - f(zn))
    return L, a, b

# Load RGB data from Excel
excel_file = 'rgb2xyz_data.xlsx'  # Replace with your Excel file name
df = pd.read_excel(excel_file)

# Extract pixel numbers and XYZ values
pixel_numbers = xyz_df['PixelNumber']
xyz_values = xyz_df[['X', 'Y', 'Z']].values

# Convert XYZ to LAB
lab_values = np.apply_along_axis(xyz_to_lab, 1, xyz_values)

# Create a new DataFrame with pixel numbers and LAB values
lab_df = pd.DataFrame(lab_values, columns=['L*', 'a*', 'b*'])
lab_df.insert(0, 'PixelNumber', pixel_numbers)

# Save LAB data to a new Excel file
lab_output_file = 'xyz2lab_data_manual.xlsx'  # Replace with your desired output file name
lab_df.to_excel(lab_output_file, index=False)
