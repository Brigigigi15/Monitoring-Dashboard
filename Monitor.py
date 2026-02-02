import pandas as pd
import matplotlib.pyplot as plt

FILE = r"C:\Users\iOne3\Desktop\DepEd\Monitoring\DEPED SWIP AND DICT ELEARNING SITES.xlsx"  # or "your_file.xls"

# Read all sheets; pick only those that have Region + Status
all_sheets = pd.read_excel(FILE, sheet_name=None)
frames = []

for name, sheet in all_sheets.items():
  if "Region" in sheet.columns and "Status" in sheet.columns:
      frames.append(sheet[["Region", "Status"]])

if not frames:
  raise ValueError("No sheet with 'Region' and 'Status' columns found.")

df = pd.concat(frames, ignore_index=True)

# Normalize
df["Region"] = df["Region"].astype(str)
df["Status"] = df["Status"].astype(str).str.strip().str.lower()

# Count per Region + Status
pivot = df.pivot_table(
  index="Region",
  columns="Status",
  aggfunc="size",
  fill_value=0,
)

# Ensure consistent order of columns
for col in ["new", "ongoing", "done"]:
  if col not in pivot.columns:
      pivot[col] = 0
pivot = pivot[["new", "ongoing", "done"]]

print("Counts per region:")
print(pivot)

# Build stacked bar chart
regions = pivot.index.tolist()
new_vals = pivot["new"].tolist()
ongoing_vals = pivot["ongoing"].tolist()
done_vals = pivot["done"].tolist()

plt.figure(figsize=(8, 4))
plt.bar(regions, new_vals, label="New")
plt.bar(regions, ongoing_vals, bottom=new_vals, label="Ongoing")
bottom_done = [n + o for n, o in zip(new_vals, ongoing_vals)]
plt.bar(regions, done_vals, bottom=bottom_done, label="Done")

plt.xlabel("Region")
plt.ylabel("Count")
plt.legend()
plt.tight_layout()
plt.show()