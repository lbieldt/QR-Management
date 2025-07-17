# Label PDF Generator

## Overview
The `label_pdf_generator.py` script generates a PDF file with images (e.g., QR codes, labels) arranged in a grid layout on A4 paper, suitable for label printing. The script supports customizable label sizes, margins, spacing, and padding, and automatically calculates right and bottom margins to ensure the layout fits within A4 dimensions (210x297 mm). It processes JPEG and PNG images from a specified folder and names the output PDF based on the first and last image filenames.

## Features
- Arranges images in a grid on A4 pages with user-defined rows and columns.
- Supports JPEG and PNG image formats.
- Automatically rotates images if a 90-degree rotation improves fit within the label area.
- Scales images to fit within padded label areas while preserving aspect ratios.
- Adds margin annotations ("Left," "Right," "Bottom") and displays layout parameters in the top margin.
- Includes first and last image filenames in the PDF for reference.
- Validates that the layout fits within A4 page dimensions.

## Requirements
- Python 3.6+
- Required Python packages:
  - `reportlab` (for PDF generation)
  - `Pillow` (for image processing)

Install the dependencies using pip:
```bash
pip install reportlab Pillow
```

## Usage
The script is executed by running `label_pdf_generator.py` with appropriate parameters. The main function, `create_label_pdf`, takes the following arguments:

- `img_folder`: Path to the folder containing JPEG or PNG images.
- `out_dir`: Directory where the output PDF will be saved.
- `lbl_width_mm`: Width of each label in millimeters.
- `lbl_height_mm`: Height of each label in millimeters.
- `lbls_x`: Number of labels per row.
- `lbls_y`: Number of labels per column.
- `mgn_left_mm`: Left margin in millimeters.
- `mgn_top_mm`: Top margin in millimeters.
- `spc_x_mm`: Horizontal spacing between labels in millimeters.
- `spc_y_mm`: Vertical spacing between labels in millimeters.
- `lbl_padding_mm`: Padding inside each label in millimeters (default: 1.0).

### Example
To generate a PDF with a grid of 8x15 labels, each 20x15 mm, with specific margins and spacing, you can run the script with the default parameters defined in the script:

```python
if __name__ == "__main__":
    IMAGE_FOLDER = "C:\\Git\\QR Codes\\qr_output"
    OUTPUT_DIR = "C:\\Git\\QR Codes"
    LABEL_WIDTH_MM = 20
    LABEL_HEIGHT_MM = 15
    LABELS_X = 8
    LABELS_Y = 15
    MARGIN_LEFT_MM = 14.5
    MARGIN_TOP_MM = 16
    SPACING_X_MM = 3
    SPACING_Y_MM = 3
    LABEL_PADDING_MM = 1
    
    create_label_pdf(
        IMAGE_FOLDER, OUTPUT_DIR, LABEL_WIDTH_MM, LABEL_HEIGHT_MM,
        LABELS_X, LABELS_Y, MARGIN_LEFT_MM, MARGIN_TOP_MM,
        SPACING_X_MM, SPACING_Y_MM, LABEL_PADDING_MM
    )
```

This will:
1. Read all JPEG/PNG images from `C:\Git\QR Codes\qr_output`.
2. Generate a PDF named after the first and last image (e.g., `firstimage_lastimage.pdf`) in `C:\Git\QR Codes`.
3. Arrange images in an 8x15 grid with 20x15 mm labels, 14.5 mm left margin, 16 mm top margin, 3 mm spacing, and 1 mm padding.

### Running the Script
1. Ensure the required dependencies are installed.
2. Place your JPEG or PNG images in the specified `img_folder`.
3. Update the parameters in the `if __name__ == "__main__":` block as needed.
4. Run the script:
   ```bash
   python label_pdf_generator.py
   ```

## Output
- The script generates a PDF file in the specified `out_dir`, named `<first_image>_<last_image>.pdf`.
- The PDF contains images arranged in a grid, with annotations for margins and layout parameters.
- Console output includes:
  - Calculated right and bottom margins.
  - List of detected images.
  - Progress messages for each page and the final PDF path.

## Notes
- Ensure the layout parameters (label size, margins, spacing) fit within A4 dimensions (210x297 mm). The script validates this and raises a `ValueError` if the layout exceeds the page size.
- Images are automatically scaled to fit within the padded label area, and rotated if necessary to optimize fit.
- The script skips PDF generation if no valid images are found in the input folder.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing
Contributions are welcome! Please submit issues or pull requests via GitLab.

## Contact
For questions or support, please open an issue in the GitLab repository.