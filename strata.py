import random
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle

def draw_table(pdf, data, x, y, col_widths):
    """Helper function to draw a table on the PDF canvas."""
    table = Table(data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    table.wrapOn(pdf, 500, 600)
    table.drawOn(pdf, x, y)



def create_pdf(filename):
    pdf = canvas.Canvas(filename, pagesize=A4)
    pdf.setTitle("Strata Tables")

    # Title
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(200, 815, "Tabulation and Observations")

    # Generate random page numbers for each stratum
    strata_pages = {
        "A-E": [random.randint(1, 316) for _ in range(16)],
        "F-J": [random.randint(317, 459) for _ in range(8)],
        "K-O": [random.randint(460, 603) for _ in range(8)],
        "P-T": [random.randint(604, 898) for _ in range(15)],
        "U-Z": [random.randint(899, 970) for _ in range(4)],
    }

    # Main Strata Table
    strata_data = [
        ["Strata", "No. of Pages", "Sample Size", "Pages No."],
        ["A-E", "316", "16", ", ".join(map(str, strata_pages["A-E"]))],
        ["F-J", "143", "8", ", ".join(map(str, strata_pages["F-J"]))],
        ["K-O", "144", "8", ", ".join(map(str, strata_pages["K-O"]))],
        ["P-T", "295", "15", ", ".join(map(str, strata_pages["P-T"]))],
        ["U-Z", "72", "4", ", ".join(map(str, strata_pages["U-Z"]))],
        ["", "970", "", ""],
    ]
    draw_table(pdf, strata_data, x=30, y=670, col_widths=[60, 65, 65, 350])

    # Individual Tables Data
    tables_data = [
        ("STRATUM A-E", [["S.No.", "Page No.", "No. of words(yi)", "yi²"]] +
         [[str(i + 1), str(page), str(words := random.randint(10, 50)), str(words**2)]
          for i, page in enumerate(strata_pages["A-E"])]),
        ("STRATUM F-J", [["S.No.", "Page No.", "No. of words(yi)", "yi²"]] +
         [[str(i + 1), str(page), str(words := random.randint(10, 50)), str(words**2)]
          for i, page in enumerate(strata_pages["F-J"])]),
        ("STRATUM K-O", [["S.No.", "Page No.", "No. of words(yi)", "yi²"]] +
         [[str(i + 1), str(page), str(words := random.randint(10, 50)), str(words**2)]
          for i, page in enumerate(strata_pages["K-O"])]),
        ("STRATUM P-T", [["S.No.", "Page No.", "No. of words(yi)", "yi²"]] +
         [[str(i + 1), str(page), str(words := random.randint(10, 50)), str(words**2)]
          for i, page in enumerate(strata_pages["P-T"])]),
        ("STRATUM U-Z", [["S.No.", "Page No.", "No. of words(yi)", "yi²"]] +
         [[str(i + 1), str(page), str(words := random.randint(10, 50)), str(words**2)]
          for i, page in enumerate(strata_pages["U-Z"])]),
    ]

    

    # Positioning tables in two columns
    x_left, x_right = 50, 300
    y_position = 345

    for i, (title, data) in enumerate(tables_data):
        pdf.setFont("Helvetica-Bold", 10)
        x = x_left if i % 2 == 0 else x_right
        if i % 2 == 0 and i != 0:
            y_position -= 180  # Move to next row after two tables

        if i == 1:
            y_position = y_position +145
        if i == 2:
            y_position = y_position -155
        if i == 3:
            y_position = y_position +12
        if i == 4:
            y_position = y_position + 50
            x = x_right
        
        # Calculate title position based on number of rows
        if i == 0:
            title_y = y_position +270+ (len(data)) * 2
        elif i == 1:
            title_y = y_position +140+ (len(data)) * 2
        elif i == 2:
            title_y = y_position + 135 + (len(data)) * 2
        elif i == 3:
            title_y = y_position + 250 + (len(data)) * 2
        else:
            title_y = y_position + 70 + (len(data)) * 2

        pdf.drawString(x, title_y, title)
        draw_table(pdf, data, x, y_position - 20, col_widths=[40, 50, 82, 50])

    # Save PDF
    pdf.save()

# Generate PDF
create_pdf("strata.pdf")