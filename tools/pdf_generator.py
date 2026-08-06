from io import BytesIO
from textwrap import wrap

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


def create_ad347_pdf(record, template_path, output_path):

    reader = PdfReader(template_path)
    writer = PdfWriter()

    page = reader.pages[0]

    width = float(page.mediabox.width)
    height = float(page.mediabox.height)

    packet = BytesIO()

    c = canvas.Canvas(packet, pagesize=(width, height))
    c.setFont("Helvetica", 10)

    # Facility
    c.drawString(110, 736, record["facility"])

    # Employee
    c.drawString(118, 709, record["employee_name"])

    # Department
    c.drawString(382, 709, record["department"])

    # Education text
    education = (
        "Our records indicate that on the following date(s): "
        + ", ".join(record["dates"])
        + ". Your timecard reflects no recorded meal breaks on the date(s) "
        "listed above. If operational needs prevent you from taking a meal "
        "break, notify your supervisor before the end of your shift so your "
        "timecard can be accurately documented."
    )

    text = c.beginText(52, 640)
    text.setFont("Helvetica", 10)

    for line in wrap(education, 88):
        text.textLine(line)

    c.drawText(text)

    c.save()

    packet.seek(0)

    overlay = PdfReader(packet)

    page.merge_page(overlay.pages[0])

    writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)
