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

    # Main Strata Table
    strata_data = [
        ["Strata", "No. of Pages", "Sample Size", "Pages No."],
        ["A-E", "316", "16", "208, 238, 266, 45, 218, 24, 227, 186, 114, 27, 39, 68, 124, 166, 326, 162"],
        ["F-J", "143", "8", "321, 382, 323, 242, 341, 348, 444, 335"],
        ["K-O", "144", "8", "472, 587, 546, 586, 646, 462, 463, 547"],
        ["P-T", "295", "15", "888, 638, 766, 853, 603, 629, 657, 678, 617, 776, 809, 834, 643, 702, 751"],
        ["U-Z", "72", "4", "926, 923, 954, 969"],
        ["", "970", "", ""],
    ]
    draw_table(pdf, strata_data, x=50, y=650, col_widths=[60, 65, 65, 350])

    # Individual Tables Data
    tables_data = [
        ("STRATUM A-E", [["S.No.", "Page No.", "No. of words(yi)", "yi²"],
                         ["1", "280", "33", "1089"], ["2", "238", "22", "484"],
                         ["3", "266", "25", "625"], ["4", "218", "28", "784"],
                         ["5", "24", "35", "1225"], ["6", "227", "22", "484"],
                         ["7", "186", "25", "625"], ["8", "124", "24", "576"],
                         ["9", "27", "39", "1521"], ["10", "68", "23", "529"]]),

        ("STRATUM F-J", [["S.No.", "Page No.", "No. of words(yi)", "yi²"],
                         ["1", "321", "16", "256"], ["2", "382", "25", "625"],
                         ["3", "323", "17", "289"], ["4", "442", "36", "1296"],
                         ["5", "341", "20", "400"], ["6", "348", "30", "900"]]),

        ("STRATUM K-O", [["S.No.", "Page No.", "No. of words(yi)", "yi²"],
                         ["1", "472", "18", "324"], ["2", "587", "17", "289"],
                         ["3", "546", "16", "256"], ["4", "586", "13", "169"],
                         ["5", "502", "21", "441"], ["6", "462", "28", "784"]] ),

        ("STRATUM P-T", [["S.No.", "Page No.", "No. of words(yi)", "yi²"],
                         ["1", "888", "31", "961"], ["2", "638", "25", "625"],
                         ["3", "766", "27", "729"], ["4", "853", "29", "841"],
                         ["5", "604", "25", "625"], ["6", "629", "27", "729"]]),

        ("STRATUM U-Z", [["S.No.", "Page No.", "No. of words(yi)", "yi²"],
                         ["1", "926", "33", "1089"], ["2", "923", "47", "2209"],
                         ["3", "954", "20", "400"], ["4", "969", "37", "1369"]]),
    ]

    # Positioning tables in two columns
    x_left, x_right = 50, 290
    y_position = 400

    for i, (title, data) in enumerate(tables_data):
        pdf.setFont("Helvetica-Bold", 10)
        x = x_left if i % 2 == 0 else x_right
        if i % 2 == 0 and i != 0:
            y_position -= 150  # Move to next row after two tables

        if(i==4):
            y_position = y_position +10
            # title_y = y_position  + (len(data))*2
        # Calculate title position based on number of rows
        if(i==0):
            title_y = y_position + 180 + (len(data))*2
        elif(i==4):
            title_y = y_position + 70 + (len(data))*2
        else:
            title_y = y_position + 80 + (len(data))*5

        pdf.drawString(x, title_y, title)
        draw_table(pdf, data, x, y_position-20, col_widths=[40, 50, 82, 50])

    # Save PDF
    pdf.save()

# Generate PDF
create_pdf("strata_tables.pdf")
