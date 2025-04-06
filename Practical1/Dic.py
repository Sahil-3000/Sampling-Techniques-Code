from openpyxl import Workbook
import random
import math
import statistics

total_pages = 550
sample_pct = 0.05
sample_size = max(1, int(round(total_pages * sample_pct)))

# Simulate dictionary: assign each page a random word count between 300 and 600
page_word_counts = {page: random.randint(7, 25) for page in range(1, total_pages+1)}

# Randomly sample pages
sample_pages = random.sample(list(page_word_counts.keys()), sample_size)
sample_counts = [page_word_counts[p] for p in sample_pages]

# Calculate estimates
avg_words = statistics.mean(sample_counts)
estimated_total = avg_words * total_pages
s = statistics.stdev(sample_counts) if sample_size > 1 else 0
std_error = (total_pages * s / math.sqrt(sample_size)) if sample_size > 1 else 0
lower_limit = estimated_total - 1.96 * std_error
upper_limit = estimated_total + 1.96 * std_error

wb = Workbook()
ws = wb.active

# Write headers for sample table and computed estimates to the right side
ws.append(["Page", "Word Count", "", "Estimated Total", "Standard Error", "95% Lower Limit", "95% Upper Limit"])
# Write sample data
for p in sample_pages:
    ws.append([p, page_word_counts[p]])
# Add computed estimates below the sample table
ws.append([])
ws.append(["", "", "", estimated_total, std_error, lower_limit, upper_limit])

wb.save("example.xlsx")
