import pandas as pd
import matplotlib.pyplot as plt

# URL of the Google Sheets document (exported as an Excel file)
excel_url = "https://docs.google.com/spreadsheets/d/13QVCFd2A1sH6lFSlJF7zW75PEtKxHmDpKNfntNds0MI/export?format=xlsx"

# Sheet name
sheet_name = "bills24"

# Read the Excel file
try:
    # Read the data into a DataFrame
    df = pd.read_excel(excel_url, sheet_name=sheet_name)
except Exception as e:
    print(f"Error reading the Excel file: {e}")
    exit()

# Select rows 2 to 18 (index 1 to 17 in Python)
df = df.iloc[1:18]

# Extract categories from Column A
categories = df.iloc[:, 0]  # Assuming the first column is Column A

# Drop rows with NaN in Column A (categories)
df = df.dropna(subset=[df.columns[0]])

# Calculate the sum of columns B to M for each row
df['Total'] = df.iloc[:, 1:13].sum(axis=1)

# Calculate the total sum of all categories
overall_total = df['Total'].sum()

# Calculate percentages for each category
df['Percentage'] = (df['Total'] / overall_total) * 100

# Combine categories, totals, and percentages into a summary DataFrame
summary_df = df[[df.columns[0], 'Total', 'Percentage']].rename(columns={df.columns[0]: "Category"})

# Display the summary
print("Category Breakdown:")
print(summary_df)

# Create the pie chart with percentages displayed outside the slices
plt.figure(figsize=(12, 8))
plt.pie(
    summary_df['Total'],
    labels=summary_df['Category'],
    autopct=lambda p: f'{p:.1f}%',
    startangle=140,
    textprops={'fontsize': 10},
    colors=plt.cm.tab20.colors,
    pctdistance=0.85,  # Position the percentage further out
    labeldistance=1.1  # Position the labels further out
)
plt.title("Category Distribution (Rows 2 to 18)", fontsize=14)
plt.show()

# Optional: Save to a new Excel file for review
output_path = "filtered_category_summary.xlsx"
summary_df.to_excel(output_path, index=False)
print(f"\nSummary saved to {output_path}")
