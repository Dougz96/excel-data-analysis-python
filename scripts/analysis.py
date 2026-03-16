import pandas as pd

# Load dataset
df = pd.read_csv("data/Sample - Superstore.csv", encoding="latin1")

print("Dataset loaded successfully\n")

# Basic dataset info
print("Dataset shape:")
print(df.shape)

print("\nColumns:")
print(df.columns)

# Total sales
total_sales = df["Sales"].sum()

print("\nTotal Sales:")
print(total_sales)

# Sales by category
sales_by_category = df.groupby("Category")["Sales"].sum()

print("\nSales by Category:")
print(sales_by_category)

# Profit by category
profit_by_category = df.groupby("Category")["Profit"].sum()

print("\nProfit by Category:")
print(profit_by_category)

# Create report
report = pd.DataFrame({
    "Sales": sales_by_category,
    "Profit": profit_by_category
})

# Save report
report.to_excel("output/report.xlsx")

print("\nReport saved to output/report.xlsx")