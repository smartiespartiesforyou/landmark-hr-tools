from io import BytesIO

from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


def wrap_text(text, font_name, font_size, max_width):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        test_line = f"{current_line} {word}".strip()

        if stringWidth(test_line, font_name, font_size) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    return lines


def create_ad347_pdf(record, template_path, output_path):
    reader = PdfReader(template_path)
    page = reader.pages[0]

    width = float(page.mediabox.width)
    height = float(page.mediabox.height)

    packet = BytesIO()
    overlay = canvas.Canvas(packet, pagesize=(width, height))

    font_name = "Helvetica"
    font_size = 10

    overlay.setFont(font_name, font_size)

    # Facility name
    overlay.drawString(
        110,
        657,
        record["facility"],
    )

    # Employee name
    overlay.drawString(
        115,
        610,
        record["employee_name"],
    )

    # Department
    overlay.drawString(
        445,
        610,
        record["department"],
    )

    # Education box boundaries
    left_x = 58
    top_y = 445
    max_width = 490
    line_height = 14

    text_object = overlay.beginText(left_x, top_y)
    text_object.setFont(font_name, font_size)
    text_object.setLeading(line_height)

    opening = (
        "Our records indicate that on the following date(s):"
    )

    for line in wrap_text(
        opening,
        font_name,
        font_size,
        max_width,
    ):
        text_object.textLine(line)

    text_object.textLine("")

    for missed_date in record["dates"]:
        text_object.textLine(f"• {missed_date}")

    text_object.textLine("")

    paragraph = (
        "Your timecard reflects no recorded meal breaks on the date(s) "
        "listed above. If operational needs prevent you from taking a meal "
        "break, notify your supervisor before the end of your shift so your "
        "timecard can be accurately documented."
    )

    for line in wrap_text(
        paragraph,
        font_name,
        font_size,
        max_width,
    ):
        text_object.textLine(line)

    overlay.drawText(text_object)
    overlay.save()

    packet.seek(0)

    overlay_pdf = PdfReader(packet)
    page.merge_page(overlay_pdf.pages[0])

    writer = PdfWriter()
    writer.add_page(page)

    with open(output_path, "wb") as output_file:
        writer.write(output_file)
