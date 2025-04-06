from openpyxl import Workbook
from openpyxl.styles import Alignment  # added import
import math
import statistics

# Experiment details
total_sheets = 600
# Given sample details
signatures = [50, 45, 38, 31, 39, 27, 18, 16, 15, 12, 10, 7, 6, 4, 3]
sheets_freq = [12, 6, 3, 2, 3, 2, 2, 1, 1, 1, 2, 2, 1, 1, 1]

# Expand the sample list based on frequencies
sample_list = []
for sign, freq in zip(signatures, sheets_freq):
    sample_list.extend([sign]*freq)
n = len(sample_list)

# Compute sample statistics
mean_sample = sum(sample_list) / n
estimated_total = total_sheets * mean_sample
s = statistics.stdev(sample_list) if n > 1 else 0
std_error = total_sheets * (s / math.sqrt(n))

# 99% confidence limits (z = 2.576)
z = 2.576
lower_limit = estimated_total - z * std_error
upper_limit = estimated_total + z * std_error

# Create and write results to an Excel workbook
wb = Workbook()
ws = wb.active

# Write sample details table and computed estimates
ws.append(["Signatures per Sheet", "No. of Sheets", "", "Estimated Total", "Std Error", "99% Lower Limit", "99% Upper Limit"])
for sign, freq in zip(signatures, sheets_freq):
    ws.append([sign, freq])
ws.append([])  # an empty row for separation
ws.append(["", "", "", estimated_total, std_error, lower_limit, upper_limit])

# Align all cells center
for row in ws.iter_rows():
    for cell in row:
        cell.alignment = Alignment(horizontal='center')

wb.save("exp2.xlsx")
