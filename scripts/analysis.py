import os
import pandas as pd
import matplotlib.pyplot as plt

# Create output folder if it does not exist
os.makedirs("output", exist_ok=True)

# Load dataset
df = pd.read_csv("data/Sample - Superstore.csv", encoding="latin1")

print("Dataset loaded successfully\n")

# Basic cleaning
df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
df = df.dropna(subset=["Order Date", "Sales", "Profit", "Category", "Sub-Category"])

print("Dataset shape after cleaning:")
print(df.shape)

print("\nColumns:")
print(df.columns.tolist())

# Total sales and profit
total_sales = df["Sales"].sum()
total_profit = df["Profit"].sum()

print(f"\nTotal Sales: {total_sales:.2f}")
print(f"Total Profit: {total_profit:.2f}")

# Sales by category
sales_by_category = df.groupby("Category")["Sales"].sum().sort_values(ascending=False)

# Profit by category
profit_by_category = df.groupby("Category")["Profit"].sum().sort_values(ascending=False)

# Top 10 sub-categories by sales
top_subcategories = (
    df.groupby("Sub-Category")["Sales"]
    .sum()
    .sort_values(ascending=False)
    .head(10)
)

# Monthly sales trend
df["YearMonth"] = df["Order Date"].dt.to_period("M").astype(str)
monthly_sales = df.groupby("YearMonth")["Sales"].sum()

print("\nSales by Category:")
print(sales_by_category)

print("\nProfit by Category:")
print(profit_by_category)

print("\nTop 10 Sub-Categories by Sales:")
print(top_subcategories)

# Export Excel report with multiple sheets
with pd.ExcelWriter("output/report.xlsx", engine="openpyxl") as writer:
    sales_by_category.reset_index(name="Sales").to_excel(
        writer, sheet_name="Sales by Category", index=False
    )
    profit_by_category.reset_index(name="Profit").to_excel(
        writer, sheet_name="Profit by Category", index=False
    )
    top_subcategories.reset_index(name="Sales").to_excel(
        writer, sheet_name="Top Sub-Categories", index=False
    )
    monthly_sales.reset_index(name="Sales").to_excel(
        writer, sheet_name="Monthly Sales", index=False
    )

# Export summary CSV
summary_df = pd.DataFrame(
    {
        "Metric": ["Total Sales", "Total Profit"],
        "Value": [total_sales, total_profit],
    }
)
summary_df.to_csv("output/summary.csv", index=False)

# Plot 1: Sales by category
plt.figure(figsize=(8, 5))
sales_by_category.plot(kind="bar")
plt.title("Sales by Category")
plt.ylabel("Sales")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("output/sales_by_category.png")
plt.close()

# Plot 2: Profit by category
plt.figure(figsize=(8, 5))
profit_by_category.plot(kind="bar")
plt.title("Profit by Category")
plt.ylabel("Profit")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("output/profit_by_category.png")
plt.close()

# Plot 3: Monthly sales trend
plt.figure(figsize=(12, 5))
monthly_sales.plot(kind="line")
plt.title("Monthly Sales Trend")
plt.ylabel("Sales")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("output/monthly_sales_trend.png")
plt.close()

print("\nFiles generated successfully in the output folder:")
print("- report.xlsx")
print("- summary.csv")
print("- sales_by_category.png")
print("- profit_by_category.png")
print("- monthly_sales_trend.png")