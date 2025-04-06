from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import random

def draw_table(pdf, data, x, y, col_widths):
    """Helper function to draw a table on the PDF canvas."""
    table = Table(data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
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
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(200, 800, "Tabulation and Observations")

    # Generate Strata Table Data
    strata_labels = ["A-E", "F-J", "K-O", "P-T", "U-Z"]
    strata_data = [["Strata", "No. of Pages", "Sample Size", "Pages No."]]
    total_pages = 0
    strata_samples = []

    for label in strata_labels:
        num_pages = random.randint(100, 400)
        sample_size = random.randint(4, 16)
        page_numbers = sorted(random.sample(range(1, num_pages + 1), sample_size))
        page_numbers_str = ", ".join(map(str, page_numbers))
        strata_data.append([label, str(num_pages), str(sample_size), page_numbers_str])
        total_pages += num_pages
        strata_samples.append(page_numbers)

    strata_data.append(["", str(total_pages), "", ""])

    # Draw main strata table
    draw_table(pdf, strata_data, x=50, y=650, col_widths=[60, 65, 65, 350])

    # Generate individual tables for each stratum
    tables_data = []

    for i, label in enumerate(strata_labels):
        title = f"STRATUM {label}"
        sample_pages = strata_samples[i]
        table_rows = [["S.No.", "Page No.", "No. of words(yi)", "yi²"]]
        for j, page_no in enumerate(sample_pages):
            word_count = random.randint(15, 50)
            word_square = word_count ** 2
            table_rows.append([str(j+1), str(page_no), str(word_count), str(word_square)])
        tables_data.append((title, table_rows))

    # Draw the individual tables in two-column layout
    x_left, x_right = 50, 290
    y_position = 400

    for i, (title, data) in enumerate(tables_data):
        pdf.setFont("Helvetica-Bold", 10)
        x = x_left if i % 2 == 0 else x_right
        if i % 2 == 0 and i != 0:
            y_position -= 150  # Move down after two tables

        # Adjust title Y based on number of rows
        if i == 0:
            title_y = y_position + 180 + (len(data)) * 2
        elif i == 4:
            title_y = y_position + 70 + (len(data)) * 2
        else:
            title_y = y_position + 80 + (len(data)) * 5

        pdf.drawString(x, title_y, title)
        draw_table(pdf, data, x, y_position - 20, col_widths=[40, 50, 82, 50])

    # Save PDF
    pdf.save()

# Generate the PDF
create_pdf("randomStrata.pdf")
