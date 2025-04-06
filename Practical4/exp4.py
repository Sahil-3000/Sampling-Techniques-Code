from openpyxl import Workbook
from openpyxl.styles import Alignment  # added import
import random
import math

# Experiment: Estimate total words using stratified random sampling.
# Total pages in dictionary.
total_pages = 970
# Sample size: max(5% of total pages, 50)
total_sample = max(int(round(total_pages * 0.05)), 50)

# Define number of strata (here, 5 equal strata)
num_strata = 5
# Calculate pages per stratum (assume equal distribution)
pages_per_stratum = total_pages // num_strata

# Simulate dictionary: word count for each page (300-600 words)
# Create list of page word counts; index 0 corresponds to page1, etc.
page_word_counts = [random.randint(7,25) for _ in range(total_pages)]

# Determine sample size per stratum using proportional allocation
sample_per_stratum = total_sample // num_strata  # here, 10 per stratum

stratum_estimates = []
stratum_vars = []
for i in range(num_strata):
    start = i * pages_per_stratum
    end = start + pages_per_stratum
    stratum_pages = list(range(start, end))
    # Stratified random sample
    sampled_indices = random.sample(stratum_pages, sample_per_stratum)
    sampled_counts = [page_word_counts[idx] for idx in sampled_indices]
    # Compute stratum mean and variance
    mean_h = sum(sampled_counts) / sample_per_stratum
    var_h = sum((x - mean_h)**2 for x in sampled_counts) / (sample_per_stratum - 1) if sample_per_stratum > 1 else 0
    stratum_estimates.append((pages_per_stratum, mean_h))
    stratum_vars.append((pages_per_stratum, sample_per_stratum, var_h))

# Estimate total words and compute overall variance from each stratum
estimated_total = sum(N_h * mean_h for N_h, mean_h in stratum_estimates)
variance_total = sum((N_h**2 * (1 - (n_h/N_h)) * var_h) / n_h for N_h, n_h, var_h in stratum_vars)
std_error = math.sqrt(variance_total)

# 99% confidence limits (z = 2.576)
z = 2.576
lower_limit = estimated_total - z * std_error
upper_limit = estimated_total + z * std_error

# Write results to Excel workbook
wb = Workbook()
ws = wb.active

# Write header indicating stratified sampling details
ws.append(["Stratum", "Pages in Stratum", "Sampled Page Indices (first 3 shown)", "", 
           "Estimated Total", "Std Error", "99% Lower Limit", "99% Upper Limit"])

# Write stratified sample details (only showing first 3 sampled indices per stratum for brevity)
for i in range(num_strata):
    start = i * pages_per_stratum
    end = start + pages_per_stratum
    stratum_pages = list(range(start+1, end+1))  # pages numbered 1...
    sampled_indices = random.sample(stratum_pages, sample_per_stratum)
    ws.append([f"Stratum {i+1}", pages_per_stratum, str(sampled_indices[:3]), "", "", "", "", ""])

# Write overall estimates below the table
ws.append([])
ws.append(["", "", "", "", estimated_total, std_error, lower_limit, upper_limit])

# Align all cells center
for row in ws.iter_rows():
    for cell in row:
        cell.alignment = Alignment(horizontal='center')

wb.save("exp3.xlsx")
