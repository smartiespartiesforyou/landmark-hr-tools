from io import BytesIO
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


def create_ad347_pdf(record, template_path, output_path):
    template = PdfReader(template_path)
    first_page = template.pages[0]

    page_width = float(first_page.mediabox.width)
    page_height = float(first_page.mediabox.height)

    overlay_stream = BytesIO()
    overlay = canvas.Canvas(
        overlay_stream,
        pagesize=(page_width, page_height)
    )

    overlay.setFont("Helvetica", 10)

    overlay.drawString(
        125,
        page_height - 108,
        record["facility"]
    )

    overlay.drawString(
        120,
        page_height - 135,
        record["employee_name"]
    )

    overlay.drawString(
        410,
        page_height - 135,
        record["department"]
    )

    education_text = (
        "Employee worked 7 or more hours without a recorded "
        "OUT and IN meal-break punch. Employee was educated "
        "to take a 30-minute uninterrupted meal break, clock "
        "OUT and IN, and notify the supervisor if work prevents "
        "the meal break."
    )

    text_box = overlay.beginText(
        55,
        page_height - 215
    )

    text_box.setFont("Helvetica", 10)

    for line in [
        education_text[0:82],
        education_text[82:164],
        education_text[164:246],
        education_text[246:]
    ]:
        text_box.textLine(line)

    overlay.drawString(
        455,
        page_height - 470,
        record["date"]
    )

    overlay.save()
    overlay_stream.seek(0)

    overlay_pdf = PdfReader(overlay_stream)
    first_page.merge_page(overlay_pdf.pages[0])

    writer = PdfWriter()
    writer.add_page(first_page)

    with open(output_path, "wb") as output_file:
        writer.write(output_file)
