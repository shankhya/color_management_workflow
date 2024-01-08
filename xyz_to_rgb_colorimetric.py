#XYZ to RGB conversion manually
import pandas as pd
import numpy as np

def gamma_correction(value):
    if value <= 0.04045:
        return 12.92 * value
    else:
        return (1.055 * (value ** (1/2.4))) - 0.055

# Ask user for the type of conversion
conversion_type = input("Enter the type of conversion (LUT/Colorimetric): ").strip().lower()

# Load the Excel file containing XYZ values
input_file = 'output_for_luts_lab_to_XYZ.xlsx'  # Replace with your file name
df = pd.read_excel(input_file, names=['Pixel', 'X', 'Y', 'Z'])

# Check user's choice and perform conversion accordingly
if conversion_type == 'lut':
    # Normalize XYZ values to the range [0, 1] for LUT-based conversion
    df[['X', 'Y', 'Z']] = df[['X', 'Y', 'Z']] / 100.0  # Assuming XYZ values are in percentage
    # Your LUT-based conversion code here...
    print("LUT-based conversion selected.")
    
elif conversion_type == 'colorimetric':
    # No normalization for colorimetric conversion
    print("Colorimetric conversion selected.")
else:
    print("Invalid input. Please choose either 'LUT' or 'Colorimetric'.")

# D65 illuminant matrix for sRGB
d65_matrix = np.array([[3.2404542, -1.5371385, -0.4985314],
                       [-0.9692660, 1.8760108, 0.0415560],
                       [0.0556434, -0.2040259, 1.0572252]])

# Convert XYZ to linear RGB
linear_rgb_values = np.dot(df[['X', 'Y', 'Z']].values, d65_matrix.T)

# Apply gamma correction for sRGB
gamma_corrected = np.vectorize(gamma_correction)
srgb_values = gamma_corrected(linear_rgb_values) * 255

# Create a new DataFrame with Pixel number and RGB values
result_df = pd.DataFrame(np.column_stack([df['Pixel'], srgb_values]),
                         columns=['Pixel', 'R', 'G', 'B'])

# Save the result to a new Excel file
output_file = 'output_xyz2rgb_data.xlsx'
result_df.to_excel(output_file, index=False)

print(f"Conversion completed and saved to {output_file}.")
