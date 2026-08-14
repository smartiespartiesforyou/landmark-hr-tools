from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


def _report_title(report_date):
    return f"Daily Attendance Report - {report_date.strftime('%B')} {report_date.day}, {report_date.year}"


def create_attendance_pdf(report_date, rows, output_path):
    pdf = canvas.Canvas(output_path, pagesize=letter)
    page_width, page_height = letter
    grouped = {}
    for row in rows:
        grouped.setdefault(row["department"], []).append(row)
    review_department = "Incomplete Punches / Still Working"
    departments = sorted(grouped, key=lambda d: (d == review_department, d))

    def page_heading(department):
        pdf.setTitle("Daily Attendance Report")
        pdf.setFont("Helvetica-Bold", 15)
        pdf.drawString(36, page_height - 45, _report_title(report_date))
        pdf.setFont("Helvetica-Bold", 12)
        pdf.drawString(36, page_height - 68, f"Department: {department}")

    def table_header(y):
        pdf.setFillColorRGB(0.12, 0.27, 0.38)
        pdf.rect(36, y - 5, page_width - 72, 22, fill=1, stroke=0)
        pdf.setFillColorRGB(1, 1, 1)
        pdf.setFont("Helvetica-Bold", 8)
        for x, label in zip((42, 205, 270, 345, 425, 485), ("Employee", "Date", "Clocked In", "Clocked Out", "Lunch", "Lunch Minutes")):
            pdf.drawString(x, y + 2, label)
        pdf.setFillColorRGB(0, 0, 0)
        return y - 24

    for department_index, department in enumerate(departments):
        if department_index:
            pdf.showPage()
        page_heading(department)
        if department == review_department:
            pdf.setFont("Helvetica", 10)
            pdf.drawString(42, page_height - 100, "These employees were excluded from the completed attendance list. Review them in UKG.")
            y = page_height - 128
            for row in grouped[department]:
                pdf.drawString(52, y, row["employee"])
                y -= 18
            continue

        y = table_header(page_height - 100)
        for row_index, row in enumerate(grouped[department]):
            if y < 65:
                pdf.showPage()
                page_heading(department)
                y = table_header(page_height - 100)
            if row_index % 2 == 0:
                pdf.setFillColorRGB(0.94, 0.96, 0.97)
                pdf.rect(36, y - 6, page_width - 72, 20, fill=1, stroke=0)
                pdf.setFillColorRGB(0, 0, 0)
            pdf.setFont("Helvetica", 8.5)
            pdf.drawString(42, y, row["employee"])
            pdf.drawString(205, y, row["date"].strftime("%m/%d/%Y"))
            pdf.drawString(270, y, row["clock_in"])
            pdf.drawString(345, y, row["clock_out"])
            pdf.drawString(425, y, row["lunch"])
            pdf.drawRightString(548, y, str(row["lunch_minutes"]))
            y -= 20
    pdf.save()
