from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


def _report_title(report_date):
    return (
        "Daily Attendance Exception Report - "
        f"{report_date.strftime('%B')} {report_date.day}, {report_date.year}"
    )


def create_attendance_pdf(report_date, rows, output_path):
    pdf = canvas.Canvas(output_path, pagesize=letter)
    page_width, page_height = letter
    grouped = {}
    for row in rows:
        grouped.setdefault(row["department"], []).append(row)

    departments = sorted(grouped)
    if not departments:
        departments = ["No Exceptions"]
        grouped["No Exceptions"] = []

    def page_heading(department):
        pdf.setTitle("Daily Attendance Exception Report")
        pdf.setFont("Helvetica-Bold", 15)
        pdf.drawString(42, page_height - 48, _report_title(report_date))
        pdf.setFont("Helvetica-Bold", 12)
        pdf.drawString(42, page_height - 72, f"Department: {department}")

    def table_header(current_y):
        pdf.setFillColorRGB(0.12, 0.27, 0.38)
        pdf.rect(42, current_y - 5, page_width - 84, 22, fill=1, stroke=0)
        pdf.setFillColorRGB(1, 1, 1)
        pdf.setFont("Helvetica-Bold", 9)
        for x, label in zip(
            (48, 240, 315, 395, 485),
            ("Employee", "Date", "Clocked In", "Clocked Out", "Lunch"),
        ):
            pdf.drawString(x, current_y + 2, label)
        pdf.setFillColorRGB(0, 0, 0)
        return current_y - 24

    for department_index, department in enumerate(departments):
        if department_index:
            pdf.showPage()

        page_heading(department)
        y = table_header(page_height - 105)
        department_rows = grouped[department]
        if not department_rows:
            pdf.setFont("Helvetica", 10)
            pdf.drawString(48, y, "No attendance exceptions found.")
        else:
            for row_index, row in enumerate(department_rows):
                if y < 70:
                    pdf.showPage()
                    page_heading(department)
                    y = table_header(page_height - 105)

                if row_index % 2 == 0:
                    pdf.setFillColorRGB(0.94, 0.96, 0.97)
                    pdf.rect(42, y - 6, page_width - 84, 20, fill=1, stroke=0)
                    pdf.setFillColorRGB(0, 0, 0)

                pdf.setFont("Helvetica", 9)
                pdf.drawString(48, y, row["employee"])
                pdf.drawString(240, y, row["date"].strftime("%m/%d/%Y"))
                pdf.drawString(315, y, row["clock_in"])
                pdf.drawString(395, y, row["clock_out"])
                pdf.drawString(485, y, row["lunch"])
                y -= 20

    pdf.save()
