# GSTR 2A Consolidator

**Version**: 1.0
**Developer**: Udit Vashisht
**Email**: [udit.vashisht@gov.in](mailto:udit.vashisht@gov.in)

## Overview

The GSTR 2A Consolidator is an application designed to consolidate monthly GSTR 2A files downloaded from GST-BO into a single annual file. The app makes it easier for businesses to compile the GSTR 2A data for an entire financial year into one spreadsheet.

## Features

- Consolidates monthly GSTR 2A Excel files into an annual file.
- Supports a user-friendly interface for file management and consolidation.
- Processes files adhering to the format required by GST-BO.
- Designed to work best with 12 files from a single financial year but can handle more.
- Outputs the consolidated file in the same directory as the input files.

## How to Use

1. Click the **'Browse Files'** button to select the monthly GSTR 2A files.
   - Ensure the files are downloaded from GST-BO.
2. The file names **must not be changed**. They should be in the format: GSTIN_MMYYYY.xlsx.
3. Although the app can consolidate more than 12 files, it is optimized for a single financial year (April to March).
4. Once you've selected the files, click **'Process Files'** to start the consolidation.
5. Click **'Clear Input'** to reset the file list.
6. The output file will be saved in the same directory where your input files are located.

## Application Interface

The home screen provides the following buttons for interaction:

- **Browse Files**: Select monthly GSTR 2A Excel files from your system.
- **Process Files**: Consolidate the selected files into an annual file.
- **Clear Input**: Clear the currently selected files.
- **Info**: View instructions on how to use the app.
- **Exit**: Close the application.

## Requirements

- Python 3.x
- Required Python packages:
- `tkinter`
- `xlsxwriter`
- `os`
- `re`
- `shutil`

## Installation

1. Clone the repository or download the source code.
2. Install the required Python packages using pip:
`pip install -r requirements.txt`

## License

This project is licensed under the MIT License.

---

For any issues or feedback, feel free to contact the developer at [udit.vashisht@gov.in](mailto:udit.vashisht@gov.in).
