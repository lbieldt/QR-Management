"""Module to generate a PDF with images arranged in a grid on A4 paper for label printing."""

import os
import io
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image

def create_label_pdf(
    img_folder: str,
    out_dir: str,
    lbl_width_mm: float,
    lbl_height_mm: float,
    lbls_x: int,
    lbls_y: int,
    mgn_left_mm: float,
    mgn_top_mm: float,
    spc_x_mm: float,
    spc_y_mm: float,
    lbl_padding_mm: float = 1.0
) -> None:
    """
    Create a PDF with images arranged in a grid on A4 paper, named after first and last images.

    Args:
        img_folder: Path to folder containing JPEG/PNG images.
        out_dir: Directory to save the output PDF.
        lbl_width_mm: Label width in millimeters.
        lbl_height_mm: Label height in millimeters.
        lbls_x: Number of labels per row.
        lbls_y: Number of labels per column.
        mgn_left_mm: Left margin in millimeters.
        mgn_top_mm: Top margin in millimeters.
        spc_x_mm: Horizontal spacing between labels in millimeters.
        spc_y_mm: Vertical spacing between labels in millimeters.
        lbl_padding_mm: Padding inside each label in millimeters (default: 1.0).
    """
    # Calculate right and bottom margins
    mgn_right_mm = 210 - mgn_left_mm - lbls_x * lbl_width_mm - (lbls_x - 1) * spc_x_mm
    mgn_bottom_mm = 297 - mgn_top_mm - lbls_y * lbl_height_mm - (lbls_y - 1) * spc_y_mm - 0.1

    # Convert mm to points (1 mm = 2.83465 points)
    label_width = lbl_width_mm * 2.83465
    label_height = lbl_height_mm * 2.83465
    margin_left = mgn_left_mm * 2.83465
    margin_right = mgn_right_mm * 2.83465
    margin_top = mgn_top_mm * 2.83465
    margin_bottom = mgn_bottom_mm * 2.83465
    spacing_x = spc_x_mm * 2.83465
    spacing_y = spc_y_mm * 2.83465
    label_padding = lbl_padding_mm * 2.83465

    # A4 page size in points
    page_width, page_height = A4

    # Validate layout fits on A4 (210x297 mm)
    total_width = margin_left + lbls_x * label_width + (lbls_x - 1) * spacing_x + margin_right
    total_height = margin_top + lbls_y * label_height + (lbls_y - 1) * spacing_y + margin_bottom
    if total_width >= 595.28 or total_height >= 841.89:
        raise ValueError(f"Layout exceeds A4 dimensions (210x297 mm): "
                         f"width={total_width/2.83465:.2f}mm, height={total_height/2.83465:.2f}mm")

    # Print calculated margins
    print(f"Calculated Margins: Right = {mgn_right_mm:.2f} mm, Bottom = {mgn_bottom_mm:.2f} mm")

    # Get list of JPEG and PNG images
    images = [f for f in os.listdir(img_folder)
              if f.lower().endswith(('.jpg', '.jpeg', '.png'))]

    # Print number of labels and their names
    print(f"Found {len(images)} label images in {img_folder}:")
    for img in images:
        print(f" - {img}")

    # Create PDF with dynamic filename
    if not images:
        print("No images found, skipping PDF generation.")
        return
    first_image_base = os.path.splitext(images[0])[0]
    last_image_base = os.path.splitext(images[-1])[0] if images else first_image_base
    out_pdf = os.path.join(out_dir, f"{first_image_base}_{last_image_base}.pdf")
    print(f"Generating PDF: {out_pdf}")
    c = canvas.Canvas(out_pdf, pagesize=A4)

    image_index = 0
    while image_index < len(images):
        # Start a new page
        print(f"Processing page for images {image_index + 1} to {min(image_index + lbls_x * lbls_y, len(images))}")
        first_image = images[image_index] if images else ""
        last_image = ""

        for y in range(lbls_y):
            for x in range(lbls_x):
                if image_index >= len(images):
                    break

                # Update last image
                last_image = images[image_index]

                # Calculate position (bottom-left corner of each label)
                x_pos = margin_left + x * (label_width + spacing_x)
                y_pos = page_height - margin_top - (y + 1) * label_height - y * spacing_y

                # Load and process image
                img_path = os.path.join(img_folder, images[image_index])
                img = Image.open(img_path)

                # Get image dimensions
                img_width, img_height = img.size
                img_aspect = img_width / img_height
                label_aspect = label_width / label_height

                # Determine if rotation improves fit
                rotate = False
                if abs(img_aspect - label_aspect) > abs((1 / img_aspect) - label_aspect):
                    img = img.rotate(90, expand=True)
                    img_width, img_height = img.size
                    rotate = True

                # Scale image to fit within padded area while maintaining aspect ratio
                padded_width = label_width - 2 * label_padding
                padded_height = label_height - 2 * label_padding
                scale = min(padded_width / img_width, padded_height / img_height)
                new_width = img_width * scale
                new_height = img_height * scale

                # Center image in the padded area
                x_offset = (label_width - new_width) / 2
                y_offset = (label_height - new_height) / 2

                # Draw image
                if rotate:
                    temp_buffer = io.BytesIO()
                    img.save(temp_buffer, format="PNG")
                    temp_buffer.seek(0)
                    img_reader = ImageReader(temp_buffer)
                    c.drawImage(img_reader, x_pos + x_offset, y_pos + y_offset,
                                new_width, new_height)
                else:
                    c.drawImage(img_path, x_pos + x_offset, y_pos + y_offset,
                                new_width, new_height)

                image_index += 1

            if image_index >= len(images):
                break

        # Draw margin text
        c.setFont("Helvetica", 10)
        c.drawCentredString(page_width / 2, margin_bottom / 2, "Bottom")

        # Draw "Left" vertically (counterclockwise)
        c.saveState()
        c.rotate(-90)
        c.drawCentredString(-page_height / 2, margin_left / 2, "Left")
        c.restoreState()

        # Draw "Right" vertically (clockwise)
        c.saveState()
        c.rotate(90)
        c.drawCentredString(page_height / 2, -(page_width - margin_right / 2), "Right")
        c.restoreState()

        # Draw first and last image filenames
        c.setFont("Helvetica", 8)
        c.drawRightString(page_width - margin_right / 2,
                          page_height - margin_top / 2 - 15, f"First: {first_image}")
        c.drawRightString(page_width - margin_right / 2,
                          margin_bottom / 2 + 5, f"Last: {last_image}")

        # Draw user-defined parameters in top margin, centered
        c.setFont("Helvetica", 4.5)
        params = [
            f"Width: {lbl_width_mm} mm, Height: {lbl_height_mm} mm, Labels X: {lbls_x}, Labels Y: {lbls_y}, "
            f"Margin Left: {mgn_left_mm} mm, Margin Right: {mgn_right_mm} mm",
            f"Margin Top: {mgn_top_mm} mm, Margin Bottom: {mgn_bottom_mm} mm, "
            f"Spacing X: {spc_x_mm} mm, Spacing Y: {spc_y_mm} mm, Padding: {lbl_padding_mm} mm"
        ]
        start_y = page_height - margin_top / 2
        for i, param in enumerate(params):
            c.drawCentredString(page_width / 2, start_y - i * 5, param)

        c.showPage()

    c.save()
    print(f"PDF generation complete: {out_pdf}")

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