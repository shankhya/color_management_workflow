#XYZ to RGB conversion manually
import pandas as pd
import numpy as np

# Load the Excel file containing XYZ values
input_file = 'output_lab_to_xyz_manual.xlsx'  # Replace with your file name
df = pd.read_excel(input_file, names=['PixelNumber', 'X', 'Y', 'Z'])

# Normalize XYZ values to the range [0, 1]
df[['X', 'Y', 'Z']] = df[['X', 'Y', 'Z']]   # Assuming XYZ values are in percentage

# D65 illuminant matrix for sRGB
d65_matrix = np.array([[3.2404542, -1.5371385, -0.4985314],
                       [-0.9692660, 1.8760108, 0.0415560],
                       [0.0556434, -0.2040259, 1.0572252]])

# Convert XYZ to linear RGB
linear_rgb_values = np.dot(df[['X', 'Y', 'Z']].values, d65_matrix.T)

# Apply gamma correction for sRGB
def gamma_correction(value):
    if value <= 0.04045:
        return 12.92 * value
    else:
        return (1.055 * (value ** (1/2.4))) - 0.055

gamma_corrected = np.vectorize(gamma_correction)
srgb_values = gamma_corrected(linear_rgb_values)*255

# Create a new DataFrame with Pixel number and RGB values
result_df = pd.DataFrame(np.column_stack([df['PixelNumber'], srgb_values]),
                         columns=['PixelNumber', 'R', 'G', 'B'])

# Save the result to a new Excel file
output_file = 'output_xyz2rgb_data_manual.xlsx'
result_df.to_excel(output_file, index=False)

