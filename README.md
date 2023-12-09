# Color Management Workflow

This repository contains a Python-based color management workflow that performs pixel-wise extraction of RGB values and conversion to CMYK, along with subsequent transformations. The workflow is divided into several steps, each implemented in a separate Python script.

## Requirements
- Python 3.x
- NumPy
- OpenCV
- pandas
- scikit-image
- scipy
- Matplotlib
- Pillow (PIL)

## Workflow Steps

### Step 1: Pixel-wise Extraction of RGB Values and Conversion to CMYK

#### Usage
```bash
python pixel_extraction_cmyk.py
```

This script reads an RGB image, performs pixel-wise extraction of RGB values, and converts them to CMYK using the Block Dye model. Light GCR and UCR calculations are applied, and the results are saved to an Excel file (`rgb_pixel_values.xlsx`).

### Step 2: Lookup Table (LUT) CMYK to Lab Calculation

#### Usage
```bash
python lut_cmyk_to_lab.py
```

This script reads a pre-defined LUT for CMYK to Lab conversion (`FOGRA39.xlsx`) and interpolates L*a*b* values based on the CMYK values obtained from the previous step. The interpolated values are saved to an Excel file (`output_for_luts.xlsx`).

### Step 3: RGB to L*a*b* Conversion

#### Usage
```bash
python rgb_to_lab.py
```

This script converts RGB values to L*a*b* using OpenCV and scikit-image libraries. The L*a*b* values are then saved to an Excel file (`lab_pixel_values.xlsx`).

### Step 4: RGB to L*a*b* Conversion (Manual)

#### Usage
```bash
python rgb_to_lab_manual.py
```

This script manually converts RGB to XYZ and then to L*a*b*. The results are saved to an Excel file (`xyz2lab_data_manual.xlsx`).

### Step 5: LAB to XYZ Conversion (Manual)

#### Usage
```bash
python lab_to_xyz_manual.py
```

This script manually converts L*a*b* to XYZ. The results are saved to an Excel file (`output_lab_to_xyz_manual.xlsx`).

### Step 6: Delta E Calculation

#### Usage
```bash
python delta_e_calculation.py
```

This script calculates Delta E values between two sets of L*a*b* values obtained from different steps. The results are saved to an Excel file (`delta_e_values_RGB_to_Lab_LUT.xlsx`), and the average Delta E is printed.

### Step 7: LUT Lab to XYZ Calculation

#### Usage
```bash
python lut_lab_to_xyz.py
```

This script reads a pre-defined LUT for Lab to XYZ conversion and interpolates XYZ values based on the L*a*b* values obtained from the previous step. The interpolated values are saved to an Excel file (`output_for_luts_lab_to_XYZ.xlsx`).

### Step 8: XYZ to RGB Conversion (Manual)

#### Usage
```bash
python xyz_to_rgb_manual.py
```

This script manually converts XYZ values to RGB. The results are saved to an Excel file (`output_xyz2rgb_data_manual.xlsx`).

### Step 9: Rearrange RGB Values to Form an Image

#### Usage
```bash
python rearrange_rgb_to_image.py
```

This script rearranges RGB values to form an image. It prompts the user to input the width and height of the image. The resulting image is saved as `output_image.png`.

### Step 10: Visualization

The visualization script (`visualization.py`) loads the original RGB image and the output image after conversion from L*a*b* to RGB. Both images are displayed side by side and saved as `output_figure.png`.

## Notes
- Make sure to replace the placeholder image paths and file names with your actual data.
- The workflow assumes that the images are in the RGB color space.

Feel free to reach out for any questions or improvements at shankhya@riptkolkata.org!

---
