#RGB to L*a*b* manually
#Part1: RGB TO XYZ
import pandas as pd
import numpy as np


def gamma_correction(rgb):
    # Apply gamma correction (inverse companding)
    gamma = 2.4
    return np.where(rgb <= 0.04045, rgb / 12.92, ((rgb + 0.055) / 1.055) ** gamma)


def rgb_to_xyz(rgb):
    # Normalize RGB values to the range of 0 to 1
    rgb_normalized = gamma_correction(rgb / 255.0)

    # RGB to XYZ transformation matrix for D65 illuminant in sRGB space
    rgb_to_xyz_matrix = np.array([
        [0.4124564, 0.3575761, 0.1804375],
        [0.2126729, 0.7151522, 0.0721750],
        [0.0193339, 0.1191920, 0.9503041]
    ])

    # Convert RGB to XYZ
    xyz = np.dot(rgb_to_xyz_matrix, rgb_normalized)
    return xyz


# Load RGB data from Excel
excel_file = 'rgb_pixel_values.xlsx'  # Replace with your Excel file name
df = pd.read_excel(excel_file)

# Extract pixel numbers and RGB values
pixel_numbers = df['PixelNumber']
rgb_values = df[['Red', 'Green', 'Blue']].values

# Convert RGB to XYZ
xyz_values = np.apply_along_axis(rgb_to_xyz, 1, rgb_values)

# Create a new DataFrame with pixel numbers and XYZ values
xyz_df = pd.DataFrame(xyz_values, columns=['X', 'Y', 'Z'])
xyz_df.insert(0, 'PixelNumber', pixel_numbers)

# Save XYZ data to a new Excel file
xyz_output_file = 'rgb2xyz_data.xlsx'  # Replace with your desired output file name
xyz_df.to_excel(xyz_output_file, index=False)
