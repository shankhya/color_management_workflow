#Rearrange RGB values to form the image
import pandas as pd
import numpy as np
from PIL import Image
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
# Load the Excel file containing modified XYZ values with headers
input_file = 'output_xyz2rgb_data_manual.xlsx'
df = pd.read_excel(input_file)
df = df.dropna()
# Extract R, G, and B values
r_values = df['R'].astype(int).values
g_values = df['G'].astype(int).values
b_values = df['B'].astype(int).values

# Assuming pixel numbers are in a column named 'Pixel'
pixel_numbers = df['PixelNumber'].astype(int).values

# Initialize matrices for R, G, and B
matrix_r = np.zeros((512, 512), dtype=np.uint8)
matrix_g = np.zeros((512, 512), dtype=np.uint8)
matrix_b = np.zeros((512, 512), dtype=np.uint8)

# Assign R, G, and B values to their respective matrices based on pixel number
matrix_r.flat[pixel_numbers - 1] = r_values
matrix_g.flat[pixel_numbers - 1] = g_values
matrix_b.flat[pixel_numbers - 1] = b_values

# Save the matrices as separate Excel sheets
with pd.ExcelWriter('output_rgb_matrices.xlsx') as writer:
    pd.DataFrame(matrix_r).to_excel(writer, sheet_name='Red', index=False, header=False)
    pd.DataFrame(matrix_g).to_excel(writer, sheet_name='Green', index=False, header=False)
    pd.DataFrame(matrix_b).to_excel(writer, sheet_name='Blue', index=False, header=False)

rgb_file = 'output_rgb_matrices.xlsx'
df_red = pd.read_excel(rgb_file, sheet_name='Red', header=None)
df_green = pd.read_excel(rgb_file, sheet_name='Green', header=None)
df_blue = pd.read_excel(rgb_file, sheet_name='Blue', header=None)

# Convert DataFrames to NumPy arrays
matrix_red = df_red.values
matrix_green = df_green.values
matrix_blue = df_blue.values

# Stack RGB matrices to create a 3D array (RGB image)
rgb_image = np.stack([matrix_red, matrix_green, matrix_blue], axis=-1)
# Convert the numpy array to uint8 (required by PIL)
rgb_image = rgb_image.astype(np.uint8)

# Create an Image object from the RGB array
image = Image.fromarray(rgb_image)

# Save the image (optional)
image.save('output_image.png')


#image.show()
im1='lena.png'
im2='output_image.png'
# Load images using matplotlib.image.imread
image1 = mpimg.imread(im1)
image2 = mpimg.imread(im2)

# Create a figure and axes
fig, axes = plt.subplots(1, 2, figsize=(10, 5))

# Display the first image with label
axes[0].imshow(image1)
axes[0].set_title('Original RGB image')

# Display the second image with label
axes[1].imshow(image2)
axes[1].set_title('Image after conversion from L*a*b* to RGB')

# Hide axes ticks
for ax in axes:
    ax.set_xticks([])
    ax.set_yticks([])

# Show the plot
plt.savefig('output_figure.png')

