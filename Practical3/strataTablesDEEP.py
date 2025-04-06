from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

# Create a PDF document with minimal margins
pdf_filename = "strata_tables_compact.pdf"
doc = SimpleDocTemplate(pdf_filename, pagesize=letter, 
                        rightMargin=36, leftMargin=36,
                        topMargin=36, bottomMargin=36)
styles = getSampleStyleSheet()
story = []

# Custom smaller font style
small_style = styles['Normal']
small_style.fontSize = 8

# Title with smaller font
title = Paragraph("Tabulation and Observations", styles['Title'])
title.style.fontSize = 12
story.append(title)
story.append(Spacer(1, 0.1*inch))

# First table data - Strain A-E
data1 = [
    ["Strain A-E", "No. of Pages", "Sample Size", "Pages No."],
    ["316", "16", "", "280, 238, 266, 452, 218, 24, 227, 186, 114, 27, 39, 68, 124, 166, 326, 162"],
    ["143", "8", "", "321, 383, 232, 442, 431, 348, 444, 335"],
    ["144", "8", "", "472, 587, 546, 586, 646, 462, 463, 547"],
    ["255", "15", "", "888, 838, 786, 853, 603, 629, 657, 628, 617, 776, 809, 834, 643, 702, 751"],
    ["72", "4", "", "926, 923, 954, 969"]
]

# Create first table with smaller column widths
t1 = Table(data1, colWidths=[0.6*inch, 0.6*inch, 0.6*inch, 3.2*inch])
t1.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
]))

story.append(t1)
story.append(Spacer(1, 0.1*inch))

# STRATUM A-E table - compact version
stratum_ae_title = Paragraph("STRATUM A-E", styles['Heading2'])
stratum_ae_title.style.fontSize = 10
story.append(stratum_ae_title)
story.append(Spacer(1, 0.05*inch))

data_ae_left = [
    ["S.No.", "Page No.", "Words(y_ij)", "v_ij"],
    ["1", "280", "33", "1089"],
    ["2", "238", "22", "484"],
    ["3", "256", "25", "625"],
    ["4", "45", "25", "625"],
    ["5", "218", "28", "784"],
    ["6", "24", "35", "1225"],
    ["7", "227", "22", "484"],
    ["8", "186", "25", "625"],
    ["9", "114", "24", "576"],
    ["10", "27", "39", "1521"],
    ["11", "39", "33", "1089"],
    ["12", "68", "23", "529"],
    ["13", "124", "18", "324"],
    ["14", "166", "24", "576"],
    ["15", "236", "20", "400"],
    ["16", "162", "24", "576"],
    ["", "", "420", "11532"]
]

data_ae_right = [
    ["S.No.", "Page No.", "Words(y_ij)", "v_ij"],
    ["1", "472", "18", "324"],
    ["2", "587", "17", "289"],
    ["3", "546", "16", "256"],
    ["4", "586", "13", "169"],
    ["5", "502", "21", "441"],
    ["6", "462", "28", "784"],
    ["7", "463", "23", "529"],
    ["8", "547", "20", "400"],
    ["", "", "156", "3192"]
]

# Combine left and right tables with a spacer column
combined_ae = []
for i in range(max(len(data_ae_left), len(data_ae_right))):
    left_row = data_ae_left[i] if i < len(data_ae_left) else ["", "", "", ""]
    right_row = data_ae_right[i] if i < len(data_ae_right) else ["", "", "", ""]
    combined_ae.append(left_row + [""] + right_row)

t_ae = Table(combined_ae, colWidths=[0.3*inch, 0.4*inch, 0.5*inch, 0.6*inch, 
                                    0.2*inch, 
                                    0.3*inch, 0.4*inch, 0.5*inch, 0.6*inch])
t_ae.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (3, 0), colors.lightgrey),
    ('BACKGROUND', (5, 0), (8, 0), colors.lightgrey),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ('SPAN', (4, 8), (4, 16)),
]))

story.append(t_ae)
story.append(Spacer(1, 0.1*inch))

# STRATUM F-J table - compact version
stratum_fj_title = Paragraph("STRATUM F-J", styles['Heading2'])
stratum_fj_title.style.fontSize = 10
story.append(stratum_fj_title)
story.append(Spacer(1, 0.05*inch))

data_fj_left = [
    ["S.No.", "Page No.", "Words(y_ij)", "v_ij"],
    ["1", "321", "16", "256"],
    ["2", "382", "25", "625"],
    ["3", "323", "17", "429"],
    ["4", "442", "36", "1296"],
    ["5", "341", "20", "400"],
    ["6", "348", "30", "900"],
    ["7", "444", "35", "1225"],
    ["8", "335", "18", "424"]
]

data_fj_right = [
    ["S.No.", "Page No.", "Words(y_ij)", "v_ij"],
    ["9", "617", "27", "729"],
    ["10", "776", "25", "625"],
    ["11", "809", "18", "324"],
    ["12", "834", "24", "576"],
    ["13", "643", "27", "729"],
    ["14", "702", "25", "625"],
    ["15", "751", "31", "961"],
    ["", "", "403", "11033"]
]

# Combine left and right tables with a spacer column
combined_fj = []
for i in range(max(len(data_fj_left), len(data_fj_right))):
    left_row = data_fj_left[i] if i < len(data_fj_left) else ["", "", "", ""]
    right_row = data_fj_right[i] if i < len(data_fj_right) else ["", "", "", ""]
    combined_fj.append(left_row + [""] + right_row)

t_fj = Table(combined_fj, colWidths=[0.3*inch, 0.4*inch, 0.5*inch, 0.6*inch, 
                                    0.2*inch, 
                                    0.3*inch, 0.4*inch, 0.5*inch, 0.6*inch])
t_fj.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (3, 0), colors.lightgrey),
    ('BACKGROUND', (5, 0), (8, 0), colors.lightgrey),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
]))

story.append(t_fj)
story.append(Spacer(1, 0.1*inch))

# STRATUM U-Z table - compact version
stratum_uz_title = Paragraph("STRATUM U-Z", styles['Heading2'])
stratum_uz_title.style.fontSize = 10
story.append(stratum_uz_title)
story.append(Spacer(1, 0.05*inch))

data_uz_left = [
    ["S.No.", "Page No.", "Words(y_ij)", "v_ij"],
    ["1", "926", "33", "1089"],
    ["2", "923", "47", "2209"],
    ["3", "954", "20", "400"],
    ["4", "969", "37", "1369"]
]

data_uz_right = [
    ["S.No.", "Page No.", "Words(y_ij)", "v_ij"],
    ["1", "926", "33", "1089"],
    ["2", "923", "47", "2209"],
    ["3", "954", "20", "400"],
    ["4", "969", "37", "1369"]
]

# Combine left and right tables with a spacer column
combined_uz = []
for i in range(max(len(data_uz_left), len(data_uz_right))):
    left_row = data_uz_left[i] if i < len(data_uz_left) else ["", "", "", ""]
    right_row = data_uz_right[i] if i < len(data_uz_right) else ["", "", "", ""]
    combined_uz.append(left_row + [""] + right_row)

t_uz = Table(combined_uz, colWidths=[0.3*inch, 0.4*inch, 0.5*inch, 0.6*inch, 
                                    0.2*inch, 
                                    0.3*inch, 0.4*inch, 0.5*inch, 0.6*inch])
t_uz.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (3, 0), colors.lightgrey),
    ('BACKGROUND', (5, 0), (8, 0), colors.lightgrey),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
]))

story.append(t_uz)
story.append(Spacer(1, 0.1*inch))

# Footer note with smaller font
footer = Paragraph('we have taken "with Neural Dictionary"', small_style)
story.append(footer)

# Build the PDF
doc.build(story)

print(f"Compact PDF created successfully: {pdf_filename}")