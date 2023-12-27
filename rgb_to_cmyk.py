#PIXEL wise extraction of RGB values and conversion to CMYK
import numpy as np
import cv2
import pandas as pd

# Read the image
image_path = 'lena.png'  # Replace with the path to your image
image = cv2.imread(image_path)
height, width, channels = image.shape
print("The width and height of the image is:", width, "X", height)
# Split the image into its RGB channels
b, g, r = cv2.split(image)

#Using Block Dye model for RGB to CMY conversion
c = ((255 - r) / 255) * 100
m = ((255 - g) / 255) * 100
y = ((255 - b) / 255) * 100

# Light GCR Calculation
k = np.minimum(c, np.minimum(m, y))
c_gcr = c - ( k)
m_gcr = m - ( k)
y_gcr = y - ( k)
k_gcr =  k

# UCR Block
# Light GCR Calculation with threshold
threshold = 200

c_ucr = np.where((c + m + y) >= threshold, (c - (0.10 * k)), c)
m_ucr = np.where((c + m + y) >= threshold, (m - (0.10 * k)), m)
y_ucr = np.where((c + m + y) >= threshold, (y - (0.10 * k)), y)
k_ucr = np.where((c + m + y) >= threshold, (0.10 * k), 0)




# Get the total number of pixels
num_pixels = image.shape[0] * image.shape[1]

# Create a list of pixel numbers
pixel_numbers = list(range(1, num_pixels + 1))

# Reshape the R, G, and B matrices to 1D arrays
r_values = r.reshape(-1)
g_values = g.reshape(-1)
b_values = b.reshape(-1)
c_values = c.reshape(-1)
m_values = m.reshape(-1)
y_values = y.reshape(-1)
k_values = k.reshape(-1)
c_ucr_values = c_ucr.reshape(-1)
m_ucr_values = m_ucr.reshape(-1)
y_ucr_values = y_ucr.reshape(-1)
k_ucr_values = k_ucr.reshape(-1)
c_gcr_values = c_gcr.reshape(-1)
m_gcr_values = m_gcr.reshape(-1)
y_gcr_values = y_gcr.reshape(-1)
k_gcr_values = k_gcr.reshape(-1)

# Create a DataFrame for the pixel numbers and RGB values
df_pixels = pd.DataFrame(pixel_numbers, columns=['PixelNumber'])
df_rgb = pd.DataFrame({'Red': r_values, 'Green': g_values, 'Blue': b_values, 'Cyan':c_values,
                       'Magenta':m_values, 'Yellow':y_values, 'Black':k_values, 'Cyan_GCR':c_gcr_values,
                       'Magenta_GCR':m_gcr_values, 'Yellow_GCR':y_gcr_values, 'Black_GCR':k_gcr_values,
                      'Cyan_UCR':c_ucr_values, 'Magenta_UCR':m_ucr_values, 'Yellow_UCR':y_ucr_values, 'Black_UCR':k_ucr_values})

# Concatenate the DataFrames
df_final = pd.concat([df_pixels, df_rgb], axis=1)

# Create a Pandas Excel writer using XlsxWriter as the engine
excel_writer = pd.ExcelWriter('rgb_pixel_values.xlsx', engine='xlsxwriter')

# Write the DataFrame to a sheet in the Excel file
df_final.to_excel(excel_writer, index=False)

# Save the Excel file
excel_writer.save()
