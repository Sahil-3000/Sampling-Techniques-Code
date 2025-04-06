from openpyxl import Workbook
from openpyxl.styles import Alignment  # added import
import random
import math

# Experiment: Estimate total words using stratified random sampling.
# Dictionary: Macmillan Essential Dictionary – ~864 pages.
# Stratified by alphabetical order: A–E, F–J, K–O, P–T, U–Z.

# Define strata distribution as (Strata, No. Of Pages)
strata_distribution = [
    ("A-E", 316),
    ("F-J", 143),
    ("K-O", 144),
    ("P-T", 295),
    ("U-Z", 72)
]
total_pages = sum(p for _, p in strata_distribution)  # 970 pages
# Use fixed sample sizes as provided: {16,8,8,15,4}
sample_sizes = [16, 8, 8, 15, 4]
total_sample = sum(sample_sizes)  # 51 pages

# Simulate dictionary: assign each page a random word count (e.g., 300-600 words)
page_word_counts = [random.randint(16,50) for _ in range(total_pages)]  # index 0 = page 1

wb = Workbook()
ws = wb.active

# --- Summary Table for Strata ---
ws.append(["Strata", "No. Of Pages", "Sample Size", "Pages No."])
current_page = 1
strata_samples = {}  # to store sampled pages for each stratum

for i, (stratum, pages) in enumerate(strata_distribution):
    sample_size = sample_sizes[i]
    pages_range = list(range(current_page, current_page + pages))
    sampled_pages = sorted(random.sample(pages_range, sample_size))
    strata_samples[stratum] = sampled_pages
    pages_str = ", ".join(map(str, sampled_pages))
    ws.append([stratum, pages, sample_size, pages_str])
    current_page += pages

ws.append([])  # blank row separation

# --- Detailed Table for Each Stratum ---
for stratum, pages in strata_distribution:
    ws.append([f"Stratum {stratum} Detail"])
    ws.append(["S.No", "Page No.", "No. Of Words", "y1j^2"])
    for i, page in enumerate(strata_samples[stratum], start=1):
        word_count = page_word_counts[page-1]  # pages numbered from 1
        ws.append([i, page, word_count, word_count**2])
    ws.append([])  # blank row after each stratum

# --- Overall Estimate ---
# (Compute overall estimate and variance using stratified sampling approach)
all_sampled_counts = []
variance_total = 0
current_page = 1
for stratum, pages in strata_distribution:
    sample_pages = strata_samples[stratum]
    sample_counts = [page_word_counts[page-1] for page in sample_pages]
    n_h = len(sample_counts)
    if n_h > 1:
        mean_h = sum(sample_counts) / n_h
        var_h = sum((x - mean_h)**2 for x in sample_counts) / (n_h - 1)
    else:
        mean_h = sample_counts[0] if sample_counts else 0
        var_h = 0
    all_sampled_counts.extend(sample_counts)
    variance_total += (pages**2 * (1 - (n_h/pages)) * var_h / n_h) if n_h > 0 else 0

overall_mean = sum(all_sampled_counts) / len(all_sampled_counts)
estimated_total = overall_mean * total_pages
std_error = math.sqrt(variance_total)
z = 2.576  # 99% confidence
lower_limit = estimated_total - z * std_error
upper_limit = estimated_total + z * std_error

ws.append([])
ws.append(["Overall Estimates"])
ws.append(["Estimated Total Words", "Std Error", "99% Lower Limit", "99% Upper Limit"])
ws.append([estimated_total, std_error, lower_limit, upper_limit])

# Align all cells center
for row in ws.iter_rows():
    for cell in row:
        cell.alignment = Alignment(horizontal='center')

wb.save("exp4.xlsx")

# --- Create PDF Report ---
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape,A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

pdf_file = "exp4.pdf"
doc = SimpleDocTemplate(
    pdf_file, 
    pagesize=A4,
    #setting margin to narrow
    leftMargin=20,
    rightMargin=20,
    topMargin=20,
    bottomMargin=20
    )
elements = []
styles = getSampleStyleSheet()

elements.append(Paragraph("Tabulation and Observations:", styles["Title"]))
# elements.append(Spacer(1, 12))

# Summary Table
summary_data = []
summary_data.append(["Strata", "No. Of Pages", "Sample Size", "Pages No."])
current_page = 1
for i, (stratum, pages) in enumerate(strata_distribution):
    sample_size = sample_sizes[i]
    sampled_pages = strata_samples[stratum]
    pages_str = ", ".join(map(str, sampled_pages))
    summary_data.append([stratum, pages, sample_size, pages_str])
    current_page += pages

summary_table = Table(summary_data)
summary_table.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,0), colors.grey),
    ("TEXTCOLOR", (0,0), (-1,0), colors.black),
    ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ("GRID", (0,0), (-1,-1), 1, colors.black)
]))
elements.append(summary_table)
# elements.append(Spacer(1, 12))

# Detailed Tables for Each Stratum
for stratum, pages in strata_distribution:
    detailed_data = []
    # Using inline markup for subscript and superscript
    detailed_data.append(["S.No", "Page No.", "No. Of Words", "y1<sub>i</sub><sup>2</sup>"])
    for i, page in enumerate(strata_samples[stratum], start=1):
        word_count = page_word_counts[page-1]
        detailed_data.append([i, page, word_count, word_count**2])
    elements.append(Paragraph(f"STRATUM {stratum}", styles["Heading2"]))
    detailed_table = Table(detailed_data)
    detailed_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.grey),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 1, colors.black)
    ]))
    elements.append(detailed_table)
    # elements.append(Spacer(1, 12))

# Overall Estimates Table
overall_data = []
overall_data.append(["Estimated Total Words", "Std Error", "99% Lower Limit", "99% Upper Limit"])
overall_data.append([estimated_total, std_error, lower_limit, upper_limit])
overall_table = Table(overall_data)
overall_table.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,0), colors.grey),
    ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
    ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ("GRID", (0,0), (-1,-1), 1, colors.black)
]))
elements.append(Paragraph("Overall Estimates", styles["Heading2"]))
elements.append(overall_table)
# elements.append(Spacer(1, 12))

doc.build(elements)



