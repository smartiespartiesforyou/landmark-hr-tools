from io import BytesIO

from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


FIELDS = {
    # Header values — positioned INSIDE the blue fields
    "facility": {
        "x": 180,
        "y": 659,
        "width": 350,
    },
    "employee": {
        "x": 195,
        "y": 611,
        "width": 205,
    },
    "department": {
        "x": 485,
        "y": 611,
        "width": 95,
    },

    # DO NOT CHANGE — education section is positioned correctly
    "education": {
        "x": 58,
        "y": 445,
        "width": 490,
    },
}


def fit_font_size(text, font_name, maximum_size, maximum_width):
    font_size = maximum_size

    while (
        font_size > 7
        and stringWidth(text, font_name, font_size) > maximum_width
    ):
        font_size -= 0.5

    return font_size


def wrap_text(text, font_name, font_size, maximum_width):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        test_line = f"{current_line} {word}".strip()

        if stringWidth(
            test_line,
            font_name,
            font_size,
        ) <= maximum_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)

            current_line = word

    if current_line:
        lines.append(current_line)

    return lines


def draw_field(pdf_canvas, text, field):
    font_name = "Helvetica"

    font_size = fit_font_size(
        text=text,
        font_name=font_name,
        maximum_size=11,
        maximum_width=field["width"],
    )

    pdf_canvas.setFont(font_name, font_size)

    pdf_canvas.drawString(
        field["x"],
        field["y"],
        text,
    )


def draw_education_box(pdf_canvas, dates):
    field = FIELDS["education"]

    font_name = "Helvetica"
    font_size = 10
    line_height = 14

    text_object = pdf_canvas.beginText(
        field["x"],
        field["y"],
    )

    text_object.setFont(font_name, font_size)
    text_object.setLeading(line_height)

    opening = "Our records indicate that on the following date(s):"

    for line in wrap_text(
        opening,
        font_name,
        font_size,
        field["width"],
    ):
        text_object.textLine(line)

    text_object.textLine("")

    for missed_date in dates:
        text_object.textLine(f"- {missed_date}")

    text_object.textLine("")

    paragraph = (
        "Your timecard reflects no recorded meal breaks on the date(s) "
        "listed above. If operational needs prevent you from taking a "
        "meal break, notify your supervisor before the end of your shift "
        "so your timecard can be accurately documented."
    )

    for line in wrap_text(
        paragraph,
        font_name,
        font_size,
        field["width"],
    ):
        text_object.textLine(line)

    pdf_canvas.drawText(text_object)


def create_ad347_pdf(record, template_path, output_path):
    reader = PdfReader(template_path)
    page = reader.pages[0]

    page_width = float(page.mediabox.width)
    page_height = float(page.mediabox.height)

    packet = BytesIO()

    overlay = canvas.Canvas(
        packet,
        pagesize=(page_width, page_height),
    )

    draw_field(
        overlay,
        record["facility"],
        FIELDS["facility"],
    )

    draw_field(
        overlay,
        record["employee_name"],
        FIELDS["employee"],
    )

    draw_field(
        overlay,
        record["department"],
        FIELDS["department"],
    )

    draw_education_box(
        overlay,
        record["dates"],
    )

    overlay.save()
    packet.seek(0)

    overlay_pdf = PdfReader(packet)
    page.merge_page(overlay_pdf.pages[0])

    writer = PdfWriter()
    writer.add_page(page)

    with open(output_path, "wb") as output_file:
        writer.write(output_file)
