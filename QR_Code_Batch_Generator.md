# QR Code Batch Generator

This project provides a Python script to generate QR codes with customizable serial numbers, saving them as PNG images and logging the details in an Excel file.

## Files

- **qr_code_generator.py**: The main Python script that generates QR codes based on user input, saves them as PNG images, and logs details in an Excel file.
- **qr_codes.xlsx**: An Excel file used to store generated QR code serials and their metadata (e.g., generation timestamp and optional notes). It may also contain a "create" sheet for inputting serials to generate.

## Features

- Generate QR codes with serial numbers in three modes:
  1. **Range Mode**: Generate QR codes from a start serial (e.g., AAA) to an end serial (e.g., AAZ).
  2. **Open-Ended Mode**: Generate a specified number of QR codes starting from a given serial (e.g., BAA).
  3. **Excel Input Mode**: Generate QR codes from serials listed in the "create" sheet of the Excel file.
- Save QR codes as PNG images with the serial number printed below each code.
- Log generated serials, timestamps, and optional notes in the "Generated" sheet of the Excel file.
- Skip duplicate serials to avoid overwriting existing QR codes.
- Configurable output directory, QR code size, font, and Excel file path.

## Prerequisites

- Python 3.6 or higher
- Required Python packages (install via pip):
  ```bash
  pip install qrcode pillow openpyxl
  ```
- A TrueType font file (e.g., Arial, located at `C:/Windows/Fonts/arial.ttf` by default) for rendering text below QR codes.
- Microsoft Excel or compatible software (optional, for viewing/editing `qr_codes.xlsx`).

## Setup

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   (Create a `requirements.txt` with `qrcode`, `pillow`, and `openpyxl` if needed.)

3. Ensure the font path in `qr_code_generator.py` (`FONT_PATH`) points to a valid TrueType font file on your system.

4. Place or create the `qr_codes.xlsx` file in the specified path (default: `C:\Git\QR Codes\qr_codes.xlsx`), or update the `EXCEL_FILE` path in the script.

## Usage

1. Run the script:
   ```bash
   python qr_code_generator.py
   ```

2. Choose a mode:
   - **Mode 1 (Range)**: Enter a start serial (e.g., `AAA`), an end serial (e.g., `AAZ`), and an optional note. Generates QR codes for all serials in the range.
   - **Mode 2 (Open-Ended)**: Enter a start serial (e.g., `BAA`), the number of QR codes to generate, and an optional note. Generates sequential QR codes.
   - **Mode 3 (Excel Input)**: Reads serials and notes from the "create" sheet in `qr_codes.xlsx`. The sheet must have headers "Serial" and "Note" in columns A and B.

3. QR codes are saved as PNG files in the `qr_output` directory (created automatically if it doesn’t exist).

4. Generated serials, timestamps, and notes are appended to the "Generated" sheet in `qr_codes.xlsx`.

### Example Excel File Structure

- **Generated Sheet** (created/updated by the script):
  | Serial | Generated At         | Note         |
  |--------|----------------------|--------------|
  | AAA    | 2025-07-17 21:38:00 | Sample note  |
  | AAB    | 2025-07-17 21:38:01 | Sample note  |

- **create Sheet** (optional, for Mode 3 input):
  | Serial | Note         |
  |--------|--------------|
  | BAA    | Custom note  |
  | BAB    | Another note |

## Configuration

Edit the following constants in `qr_code_generator.py` to customize the script:
- `OUTPUT_DIR`: Directory to save QR code PNGs (default: `qr_output`).
- `EXCEL_FILE`: Path to the Excel file (default: `C:\Git\QR Codes\qr_codes.xlsx`).
- `FONT_PATH`: Path to the TrueType font file (default: `C:/Windows/Fonts/arial.ttf`).
- `FONT_SIZE`: Font size for text below QR codes (default: `80`).
- `QR_SIZE`: Size of the QR code in pixels (default: `300`).

## Notes

- The script checks for existing serials in the "Generated" sheet to avoid duplicates.
- If the font file is not found, it falls back to a default font (which may not render text as expected).
- Ensure the Excel file path is accessible and writable to avoid errors.
- Mode 2 has a safety limit (`max_attempts = 10000`) to prevent infinite loops if too many duplicates are encountered.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please submit a pull request or open an issue for bug reports or feature requests.