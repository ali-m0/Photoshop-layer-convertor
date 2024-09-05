# Convert and Map Excel File Cells to Photoshop Layers and Export as Images

This project is a desktop application built with the tkinter library that allows users to map each column from an Excel or CSV file to specific Photoshop layers and convert them directly into images. Unlike the standard method in Photoshop, which requires a CSV or text file with a custom column format to create datasets, this application enables users to use any CSV or Excel file with multiple columns.

## Features

- **User-Friendly Interface**: Built with tkinter, providing an easy-to-use graphical interface for file selection and mapping configuration.
- **Flexible Input Files**: Supports both CSV and Excel files with any number of columns.
- **Custom Mapping**: Users can map each column in the file to specific Photoshop layers.
- **Automated Image Export**: Automatically generates and exports images directly from Photoshop after mapping.

## Problem Solved

Previously, creating datasets in Photoshop required a specific format for CSV or text files. Users had to manually format their data files to match these requirements, limiting flexibility and increasing the chances of errors. This project provides a solution by allowing any CSV or Excel file to be used directly, thereby saving time and effort.

## How It Works

1. **Input File**: Load your CSV or Excel file containing the data you want to map to Photoshop layers through the application's graphical interface.
2. **Column to Layer Mapping**: Use the interface to define how each column in your input file should map to a specific layer in Photoshop.
3. **Automated Processing**: The application processes each row of data, updates the mapped layers in Photoshop accordingly, and exports the result as an image.

## Installation

To run this project, you'll need:
- Python 3.x
- Libraries: pandas, openpyxl (for Excel file handling), photoshop-python-api, tkinter (comes pre-installed with Python)
- Photoshop with scripting support enabled

```bash```
pip install pandas openpyxl photoshop-python-api

## Usage

Prepare your input file and run the application using these steps:

```bash```
# Prepare your input file:
# Create or use an existing CSV or Excel file containing the data to be mapped.

# Run the application:
python convert_and_map_excel_to_photoshop.py

# Select input file:
# Use the tkinter interface to browse and select your CSV or Excel file.

# Define mappings:
# Use the application interface to define how each column should map to Photoshop layers.

# Process and export:
# Click the "Process" button to map the data and export the images.

## Example

Consider a CSV file with the following columns:

| Name  | Age | Background Color | Font Size |
|-------|-----|------------------|-----------|
| John  | 28  | Red              | 12pt      |
| Alice | 30  | Blue             | 14pt      |

You can map these columns to Photoshop layers (e.g., text layers, background layers, etc.) using the application, and it will generate images with the specified data.

## Known Issues

- **Performance**: Processing large files may be slow due to the limitations of Photoshop scripting.
- **Photoshop Compatibility**: Ensure your version of Photoshop supports Python scripting.

## Contribution

If you have suggestions or want to contribute to this project, please contact me directly.

### Contact Information:

- **Author**: Ali Mansouri
- **LinkedIn**: [Ali Mansouri](https://www.linkedin.com/in/ali-mansouri-a7984215b/)
- **Email**: [ali.mansouri1998@gmail.com](mailto:ali.mansouri1998@gmail.com)

## License

For licensing information, please contact the author directly at the email provided above.