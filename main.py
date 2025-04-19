import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from datetime import datetime
from typing import List

# Global variable to track image titles
_mfajlsdf98q21_image_title_list = []

# Set style for the plots
plt.style.use('seaborn')

# Function to save and display the plot
def save_and_show(title):
    global _mfajlsdf98q21_image_title_list
    plt.title(title)
    _mfajlsdf98q21_image_title_list.append(title)
    plt.tight_layout()
    plt.show()

# Load the data
print("Loading Superstore dataset...")
df = pd.read_excel('Superstore.xlsx')

# Display basic information
print("Dataset Overview:")
print(f"Shape: {df.shape}")
print(f"Columns: {df.columns.tolist()}")
print("\nFirst few rows:")
print(df.head())

# Data cleaning
print("\nChecking for missing values...")
missing_values = df.isnull().sum()
print(missing_values[missing_values > 0] if any(missing_values > 0) else "No missing values found.")

# Convert dates to datetime if they aren't already
if not pd.api.types.is_datetime64_any_dtype(df['Order Date']):
    df['Order Date'] = pd.to_datetime(df['Order Date'])
if not pd.api.types.is_datetime64_any_dtype(df['Ship Date']):
    df['Ship Date'] = pd.to_datetime(df['Ship Date'])

# Add year, month columns for easier analysis
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.month
df['MonthYear'] = df['Order Date'].dt.to_period('M')

# Basic statistics
print("\nBasic statistics of Sales and Profit:")
print(df[['Sales', 'Profit']].describe())

# 1. Sales Trend Over Time
print("\nCreating sales trend visualization...")
monthly_sales = df.groupby('MonthYear')['Sales'].sum().reset_index()
monthly_sales['MonthYear'] = monthly_sales['MonthYear'].astype(str)

plt.figure(figsize=(12, 6))
plt.plot(monthly_sales['MonthYear'], monthly_sales['Sales'], marker='o', linewidth=2)
plt.xticks(rotation=45)
plt.xlabel('Month-Year')
plt.ylabel('Total Sales ($)')
plt.grid(True, alpha=0.3)
save_and_show('Monthly Sales Trend')

# 2. Sales by Region
print("\nAnalyzing sales by region...")
region_sales = df.groupby('Region')['Sales'].sum().sort_values(ascending=False)

plt.figure(figsize=(10, 6))
bars = plt.bar(region_sales.index, region_sales.values, color='skyblue')
plt.xlabel('Region')
plt.ylabel('Total Sales ($)')
plt.grid(True, alpha=0.3, axis='y')

# Add values on top of bars
for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2., height + 5000,
             f'${height:,.0f}', ha='center', va='bottom', rotation=0)

save_and_show('Sales by Region')

# 3. Profitability by Category
print("\nAnalyzing profitability by category...")
category_profit = df.groupby('Category')[['Sales', 'Profit']].sum()
category_profit['Profit_Margin'] = (category_profit['Profit'] / category_profit['Sales']) * 100

plt.figure(figsize=(10, 6))
bars = plt.bar(category_profit.index, category_profit['Profit'], color='lightgreen')
plt.xlabel('Category')
plt.ylabel('Total Profit ($)')
plt.grid(True, alpha=0.3, axis='y')

# Add profit values and margin percentages on top of bars
for bar in bars:
    height = bar.get_height()
    category = bar.get_x() + bar.get_width()/2.
    profit_margin = category_profit.loc[category_profit.index[bars.index(bar)]]['Profit_Margin']
    plt.text(category, height + 5000,
             f'${height:,.0f}\n({profit_margin:.1f}%)', ha='center', va='bottom')

save_and_show('Profit by Category')

# 4. Sales by Customer Segment
print("\nAnalyzing sales by customer segment...")
segment_sales = df.groupby('Segment')['Sales'].sum().sort_values(ascending=False)

plt.figure(figsize=(10, 6))
plt.pie(segment_sales, labels=segment_sales.index, autopct='%1.1f%%', 
        startangle=90, shadow=False, explode=[0.05]*len(segment_sales),
        colors=['#ff9999','#66b3ff','#99ff99'])
plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
save_and_show('Sales by Customer Segment')

# 5. Sub-Category Performance
print("\nAnalyzing sub-category performance...")
subcat_performance = df.groupby('Sub-Category').agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Order ID': pd.Series.nunique  # Count unique orders as a measure of popularity
}).sort_values('Sales', ascending=False)

subcat_performance = subcat_performance.rename(columns={'Order ID': 'Order Count'})
subcat_performance['Profit_Margin'] = (subcat_performance['Profit'] / subcat_performance['Sales']) * 100

# Top 10 sub-categories by sales
top_subcats = subcat_performance.head(10)

plt.figure(figsize=(12, 8))
bars = plt.barh(top_subcats.index[::-1], top_subcats['Sales'][::-1], color='lightblue')
plt.xlabel('Total Sales ($)')
plt.ylabel('Sub-Category')
plt.grid(True, alpha=0.3, axis='x')

# Add values to bars
for bar in bars:
    width = bar.get_width()
    plt.text(width + 5000, bar.get_y() + bar.get_height()/2.,
             f'${width:,.0f}', ha='left', va='center')

save_and_show('Top 10 Sub-Categories by Sales')

# 6. Profit Margin by Sub-Category
print("\nAnalyzing profit margin by sub-category...")
# Sort by profit margin
profit_margin_subcat = subcat_performance.sort_values('Profit_Margin', ascending=False)

plt.figure(figsize=(12, 8))
bars = plt.barh(profit_margin_subcat.index[::-1], profit_margin_subcat['Profit_Margin'][::-1])

# Color bars based on positive/negative profit
colors = ['green' if x >= 0 else 'red' for x in profit_margin_subcat['Profit_Margin'][::-1]]
for bar, color in zip(bars, colors):
    bar.set_color(color)

plt.xlabel('Profit Margin (%)')
plt.ylabel('Sub-Category')
plt.axvline(x=0, color='black', linestyle='-', alpha=0.3)
plt.grid(True, alpha=0.3, axis='x')

# Add values to bars
for bar in bars:
    width = bar.get_width()
    plt.text(width + 1 if width >= 0 else width - 1, 
             bar.get_y() + bar.get_height()/2.,
             f'{width:.1f}%', ha='left' if width >= 0 else 'right', va='center')

save_and_show('Profit Margin by Sub-Category')

# 7. Discount Analysis
print("\nAnalyzing the impact of discounts on profit...")
# Create discount bins
df['Discount_Bin'] = pd.cut(df['Discount'], 
                           bins=[0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 1.0],
                           labels=['0-10%', '10-20%', '20-30%', '30-40%', '40-50%', 
                                  '50-60%', '60-70%', '70-80%', '80-100%'])

discount_analysis = df.groupby('Discount_Bin').agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Order ID': 'count'
}).reset_index()

discount_analysis['Profit_Margin'] = (discount_analysis['Profit'] / discount_analysis['Sales']) * 100
discount_analysis = discount_analysis.rename(columns={'Order ID': 'Number of Orders'})

fig, ax1 = plt.subplots(figsize=(12, 6))

color = 'tab:blue'
ax1.set_xlabel('Discount Level')
ax1.set_ylabel('Number of Orders', color=color)
ax1.bar(discount_analysis['Discount_Bin'], discount_analysis['Number of Orders'], color=color, alpha=0.7)
ax1.tick_params(axis='y', labelcolor=color)

ax2 = ax1.twinx()
color = 'tab:red'
ax2.set_ylabel('Profit Margin (%)', color=color)
ax2.plot(discount_analysis['Discount_Bin'], discount_analysis['Profit_Margin'], color=color, marker='o', linewidth=2)
ax2.tick_params(axis='y', labelcolor=color)

plt.xticks(rotation=45)
plt.grid(False)
save_and_show('Discount Analysis - Orders vs. Profit Margin')

# 8. Shipping Performance
print("\nAnalyzing shipping performance...")
# Calculate shipping days
df['Shipping_Days'] = (df['Ship Date'] - df['Order Date']).dt.days

# Average shipping days by ship mode
shipping_performance = df.groupby('Ship Mode').agg({
    'Shipping_Days': 'mean',
    'Order ID': pd.Series.nunique
}).sort_values('Shipping_Days')

shipping_performance = shipping_performance.rename(columns={'Order ID': 'Number of Orders'})

fig, ax1 = plt.subplots(figsize=(10, 6))

color = 'tab:green'
ax1.set_xlabel('Ship Mode')
ax1.set_ylabel('Average Shipping Days', color=color)
ax1.bar(shipping_performance.index, shipping_performance['Shipping_Days'], color=color, alpha=0.7)
ax1.tick_params(axis='y', labelcolor=color)

ax2 = ax1.twinx()
color = 'tab:purple'
ax2.set_ylabel('Number of Orders', color=color)
ax2.plot(shipping_performance.index, shipping_performance['Number of Orders'], color=color, marker='s', linewidth=2)
ax2.tick_params(axis='y', labelcolor=color)

plt.grid(False)
save_and_show('Shipping Performance by Ship Mode')

# Summary of findings
print("\nSummary of Key Findings:")
print(f"1. Total Sales: ${df['Sales'].sum():,.2f}")
print(f"2. Total Profit: ${df['Profit'].sum():,.2f}")
print(f"3. Overall Profit Margin: {(df['Profit'].sum() / df['Sales'].sum()) * 100:.2f}%")
print(f"4. Best Performing Region: {region_sales.index[0]} (${region_sales.values[0]:,.2f} in sales)")
print(f"5. Most Profitable Category: {category_profit['Profit'].idxmax()} (${category_profit['Profit'].max():,.2f})")
print(f"6. Largest Customer Segment: {segment_sales.index[0]} (${segment_sales.values[0]:,.2f} in sales)")
print(f"7. Most Profitable Sub-Category: {profit_margin_subcat.index[0]} ({profit_margin_subcat['Profit_Margin'][0]:.2f}% margin)")
print(f"8. Least Profitable Sub-Category: {profit_margin_subcat.index[-1]} ({profit_margin_subcat['Profit_Margin'][-1]:.2f}% margin)")

print("\nComplete analysis finished. All visualizations generated.")
