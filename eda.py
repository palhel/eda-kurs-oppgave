
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

df_sales_orders = pd.read_excel('Regional Sales Dataset.xlsx', sheet_name='Sales Orders')
df_customers = pd.read_excel('Regional Sales Dataset.xlsx', sheet_name='Customers')
df_products = pd.read_excel('Regional Sales Dataset.xlsx', sheet_name='Products')
df_regions = pd.read_excel('Regional Sales Dataset.xlsx', sheet_name='Regions')
df_state_regions = pd.read_excel('Regional Sales Dataset.xlsx', sheet_name='State Regions', header=1)
df_2017_budget = pd.read_excel('Regional Sales Dataset.xlsx', sheet_name='2017 Budgets')

# Check for missing data
print("\nMissing values per column:")
print(df_sales_orders.isnull().sum())
print(df_customers.isnull().sum())
print(df_products.isnull().sum())
print(df_regions.isnull().sum())
print(df_state_regions.isnull().sum())
print(df_2017_budget.isnull().sum())

#understand the Sales order data
print("Sales Orders data:")
print(df_sales_orders.head())

print("\nRegions data:")
print(df_regions.head())

print("\nState Regions data:")
print(df_state_regions.head())


# Step 1: Join Sales Orders with Regions to get state
print("\n=== STEP 1: Join Sales Orders with Regions ===")
sales_with_states = df_sales_orders.merge(df_regions, left_on='Delivery Region Index', right_on='id', how='left')
print("Sales data now includes state information")
print(sales_with_states[['Delivery Region Index', 'state', 'Line Total']].head())

# Step 2: Join with State Regions to get geographic region
print("\n=== STEP 2: Join with State Regions to get geographic region ===")
print("State Regions columns:", df_state_regions.columns.tolist())

# The State Regions data has 'State' and 'Region' columns (not 'state' and 'region')
sales_with_regions = sales_with_states.merge(df_state_regions, left_on='state', right_on='State', how='left')
print("Sales data now includes geographic region")
print(sales_with_regions[['Delivery Region Index', 'state', 'Region', 'Line Total']].head())

# Step 3: Calculate total sales by geographic region
print("\n=== STEP 3: Calculate sales by geographic region ===")
sales_by_geographic_region = sales_with_regions.groupby('Region')['Line Total'].sum().sort_values(ascending=False)
print("Total sales by geographic region:")
print(sales_by_geographic_region)


# Step 4: Create graph

plt.figure(figsize=(10, 6))
sns.barplot(x=sales_by_geographic_region.index, y=sales_by_geographic_region.values, palette='viridis')
plt.title('Total Sales by Geographic Region', fontsize=16, fontweight='bold')
plt.xlabel('Geographic Region', fontsize=12)
plt.ylabel('Total Sales ($)', fontsize=12)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('regional_sales.png', dpi=300, bbox_inches='tight')
plt.show()



# Step 5: Sales by Year AND Region

# First, let's see what years we have
sales_with_regions['OrderDate'] = pd.to_datetime(sales_with_regions['OrderDate'])
sales_with_regions['Year'] = sales_with_regions['OrderDate'].dt.year

print("Years in dataset:", sorted(sales_with_regions['Year'].unique()))

# Calculate sales by year and region
sales_by_year_region = sales_with_regions.groupby(['Year', 'Region'])['Line Total'].sum().reset_index()
print("\nSales by Year and Region:")
print(sales_by_year_region)
# Step 6: More Meaningful Business Metrics
print("\n=== STEP 6: Revenue per Household and Profit Analysis ===")

# Calculate profit (Revenue - Cost)
sales_with_regions['Profit'] = sales_with_regions['Line Total'] - sales_with_regions['Total Unit Cost']

# Calculate revenue per household by region
print("Calculating revenue per household by region...")

# Group by region and calculate metrics
region_metrics = sales_with_regions.groupby('Region').agg({
    'Line Total': 'sum',           # Total revenue
    'Profit': 'sum',               # Total profit
    'households': 'first'          # Households (should be same for all records in a region)
}).reset_index()

# Calculate revenue per household
region_metrics['Revenue_per_Household'] = region_metrics['Line Total'] / region_metrics['households']
region_metrics['Profit_per_Household'] = region_metrics['Profit'] / region_metrics['households']

print("\n📊 Business Metrics by Region:")
print(region_metrics[['Region', 'Line Total', 'Profit', 'households', 'Revenue_per_Household', 'Profit_per_Household']].round(2))

# Create visualization for revenue per household
plt.figure(figsize=(12, 5))

plt.subplot(1, 2, 1)
sns.barplot(data=region_metrics, x='Region', y='Revenue_per_Household', palette='viridis')
plt.title('Revenue per Household by Region', fontweight='bold')
plt.ylabel('Revenue per Household ($)')
plt.xticks(rotation=45)

plt.subplot(1, 2, 2)
sns.barplot(data=region_metrics, x='Region', y='Profit_per_Household', palette='viridis')
plt.title('Profit per Household by Region', fontweight='bold')
plt.ylabel('Profit per Household ($)')
plt.xticks(rotation=45)

plt.tight_layout()
plt.savefig('household_metrics.png', dpi=300, bbox_inches='tight')
plt.show()

print("📊 Household metrics chart saved as 'household_metrics.png'")

# Step 7: Profit per Sale Analysis
print("\n=== STEP 7: Profit per Sale by Region ===")

# Calculate profit per sale by region
profit_per_sale = sales_with_regions.groupby('Region').agg({
    'Profit': 'mean',              # Average profit per sale
    'Line Total': 'mean',          # Average revenue per sale
    'Total Unit Cost': 'mean'      # Average cost per sale
}).reset_index()

profit_per_sale['Profit_Margin_Percent'] = (profit_per_sale['Profit'] / profit_per_sale['Line Total']) * 100

print("\n📊 Profit per Sale Analysis:")
print(profit_per_sale[['Region', 'Profit', 'Line Total', 'Total Unit Cost', 'Profit_Margin_Percent']].round(2))

# Create visualization for profit per sale
plt.figure(figsize=(15, 5))

plt.subplot(1, 3, 1)
sns.barplot(data=profit_per_sale, x='Region', y='Profit', palette='viridis')
plt.title('Average Profit per Sale by Region', fontweight='bold')
plt.ylabel('Average Profit per Sale ($)')
plt.xticks(rotation=45)

plt.subplot(1, 3, 2)
sns.barplot(data=profit_per_sale, x='Region', y='Line Total', palette='viridis')
plt.title('Average Revenue per Sale by Region', fontweight='bold')
plt.ylabel('Average Revenue per Sale ($)')
plt.xticks(rotation=45)

plt.subplot(1, 3, 3)
sns.barplot(data=profit_per_sale, x='Region', y='Profit_Margin_Percent', palette='viridis')
plt.title('Profit Margin % by Region', fontweight='bold')
plt.ylabel('Profit Margin (%)')
plt.xticks(rotation=45)

plt.tight_layout()
plt.savefig('profit_per_sale.png', dpi=300, bbox_inches='tight')
plt.show()

print("📊 Profit per sale chart saved as 'profit_per_sale.png'")

# Find the most profitable region per sale
best_profit_region = profit_per_sale.loc[profit_per_sale['Profit'].idxmax(), 'Region']
best_profit_value = profit_per_sale.loc[profit_per_sale['Profit'].idxmax(), 'Profit']
print(f"\n🎯 BEST PERFORMER: {best_profit_region} with ${best_profit_value:.2f} average profit per sale")

# Step 8: Channel Analysis
print("\n=== STEP 8: Sales Channel Analysis ===")

# First, let's see what channels we have
print("Available sales channels:")
print(sales_with_regions['Channel'].value_counts())

# Channel performance by region
print("\n📊 Channel Performance by Region:")

# Create a pivot table: Region vs Channel
channel_by_region = sales_with_regions.groupby(['Region', 'Channel'])['Line Total'].sum().unstack(fill_value=0)
print("\nTotal Sales by Region and Channel:")
print(channel_by_region)

# Calculate channel percentages by region
channel_percentages = channel_by_region.div(channel_by_region.sum(axis=1), axis=0) * 100
print("\nChannel Distribution by Region (%):")
print(channel_percentages.round(1))

# Create visualization
plt.figure(figsize=(15, 10))

# Chart 1: Total sales by channel and region
plt.subplot(2, 2, 1)
channel_by_region.plot(kind='bar', ax=plt.gca(), color=['#1f77b4', '#ff7f0e', '#2ca02c'])
plt.title('Total Sales by Region and Channel', fontweight='bold')
plt.ylabel('Total Sales ($)')
plt.xlabel('Region')
plt.legend(title='Channel')
plt.xticks(rotation=45)

# Chart 2: Channel distribution percentages
plt.subplot(2, 2, 2)
channel_percentages.plot(kind='bar', ax=plt.gca(), color=['#1f77b4', '#ff7f0e', '#2ca02c'])
plt.title('Channel Distribution by Region (%)', fontweight='bold')
plt.ylabel('Percentage (%)')
plt.xlabel('Region')
plt.legend(title='Channel')
plt.xticks(rotation=45)

# Chart 3: Average profit by channel
plt.subplot(2, 2, 3)
profit_by_channel = sales_with_regions.groupby('Channel')['Profit'].mean()
sns.barplot(x=profit_by_channel.index, y=profit_by_channel.values, palette='viridis')
plt.title('Average Profit per Sale by Channel', fontweight='bold')
plt.ylabel('Average Profit per Sale ($)')
plt.xticks(rotation=45)

# Chart 4: Channel performance by region (heatmap)
plt.subplot(2, 2, 4)
sns.heatmap(channel_by_region, annot=True, fmt='.0f', cmap='YlOrRd')
plt.title('Sales Heatmap: Region vs Channel', fontweight='bold')
plt.xlabel('Channel')
plt.ylabel('Region')

plt.tight_layout()
plt.savefig('channel_analysis.png', dpi=300, bbox_inches='tight')
plt.show()


# Step 9: Product Performance Analysis
print("\n=== STEP 9: Product Performance Analysis ===")

# First, let's see what products we have
print("Available products (first 10):")
print(sales_with_regions['Product Description Index'].value_counts().head(10))

# Join with products data to get product names
print("\nJoining with product names...")
sales_with_products = sales_with_regions.merge(df_products, left_on='Product Description Index', right_on='Index', how='left')

print("Products data:")
print(df_products.head())

# Top products by total sales
print("\n📊 Top 10 Products by Total Sales:")
top_products = sales_with_products.groupby('Product Name')['Line Total'].sum().sort_values(ascending=False).head(10)
print(top_products)

# Top products by profit
print("\n📊 Top 10 Products by Total Profit:")
top_products_profit = sales_with_products.groupby('Product Name')['Profit'].sum().sort_values(ascending=False).head(10)
print(top_products_profit)

# Product performance by region
print("\n📊 Product Performance by Region:")

# Create visualization
plt.figure(figsize=(20, 15))

# Chart 1: Top 10 products by sales
plt.subplot(2, 3, 1)
top_products.plot(kind='bar', color='skyblue')
plt.title('Top 10 Products by Total Sales', fontweight='bold')
plt.ylabel('Total Sales ($)')
plt.xticks(rotation=45, ha='right')

# Chart 2: Top 10 products by profit
plt.subplot(2, 3, 2)
top_products_profit.plot(kind='bar', color='lightcoral')
plt.title('Top 10 Products by Total Profit', fontweight='bold')
plt.ylabel('Total Profit ($)')
plt.xticks(rotation=45, ha='right')

# Chart 3: Average profit margin by product
plt.subplot(2, 3, 3)
product_margins = sales_with_products.groupby('Product Name').agg({
    'Profit': 'mean',
    'Line Total': 'mean'
}).reset_index()
product_margins['Profit_Margin_Percent'] = (product_margins['Profit'] / product_margins['Line Total']) * 100
top_margin_products = product_margins.nlargest(10, 'Profit_Margin_Percent')
sns.barplot(data=top_margin_products, x='Profit_Margin_Percent', y='Product Name', palette='viridis')
plt.title('Top 10 Products by Profit Margin %', fontweight='bold')
plt.xlabel('Profit Margin (%)')

# Chart 4: Product sales by region (heatmap for top products)
plt.subplot(2, 3, 4)
top_5_products = top_products.head(5).index
product_region_sales = sales_with_products[sales_with_products['Product Name'].isin(top_5_products)].groupby(['Product Name', 'Region'])['Line Total'].sum().unstack(fill_value=0)
sns.heatmap(product_region_sales, annot=True, fmt='.0f', cmap='YlOrRd')
plt.title('Top 5 Products: Sales by Region', fontweight='bold')
plt.xlabel('Region')
plt.ylabel('Product')

# Chart 5: Product frequency (how often each product is sold)
plt.subplot(2, 3, 5)
product_frequency = sales_with_products['Product Name'].value_counts().head(10)
product_frequency.plot(kind='bar', color='lightgreen')
plt.title('Top 10 Most Frequently Sold Products', fontweight='bold')
plt.ylabel('Number of Sales')
plt.xticks(rotation=45, ha='right')

# Chart 6: Average sale value by product
plt.subplot(2, 3, 6)
avg_sale_by_product = sales_with_products.groupby('Product Name')['Line Total'].mean().sort_values(ascending=False).head(10)
avg_sale_by_product.plot(kind='bar', color='gold')
plt.title('Top 10 Products by Average Sale Value', fontweight='bold')
plt.ylabel('Average Sale Value ($)')
plt.xticks(rotation=45, ha='right')

plt.tight_layout()
plt.savefig('product_analysis.png', dpi=300, bbox_inches='tight')
plt.show()


# Find the most profitable product
best_product = top_products_profit.index[0]
best_product_profit = top_products_profit.iloc[0]
print(f"\n🎯 MOST PROFITABLE PRODUCT: {best_product} with ${best_product_profit:,.2f} total profit")

# Step 10: Seasonal Analysis (REQUIRED for assignment)
print("\n=== STEP 10: Seasonal Analysis ===")

# Extract month and quarter from dates
sales_with_products['Month'] = sales_with_products['OrderDate'].dt.month
sales_with_products['Quarter'] = sales_with_products['OrderDate'].dt.quarter
sales_with_products['Month_Name'] = sales_with_products['OrderDate'].dt.month_name()

print("Analyzing seasonal patterns...")

# Seasonal sales by month
monthly_sales = sales_with_products.groupby('Month_Name')['Line Total'].sum().reindex([
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
])

print("\n📊 Sales by Month:")
print(monthly_sales)

# Seasonal sales by quarter
quarterly_sales = sales_with_products.groupby('Quarter')['Line Total'].sum()
print("\n📊 Sales by Quarter:")
print(quarterly_sales)

# Seasonal analysis by region
# Create seasonal visualizations
plt.figure(figsize=(20, 12))

# Chart 1: Monthly sales trend
plt.subplot(2, 3, 1)
monthly_sales.plot(kind='line', marker='o', color='blue', linewidth=2)
plt.title('Monthly Sales Trend', fontweight='bold', fontsize=14)
plt.ylabel('Total Sales ($)')
plt.xlabel('Month')
plt.xticks(rotation=45)
plt.grid(True, alpha=0.3)

# Chart 2: Quarterly sales
plt.subplot(2, 3, 2)
quarterly_sales.plot(kind='bar', color='green', alpha=0.7)
plt.title('Quarterly Sales', fontweight='bold', fontsize=14)
plt.ylabel('Total Sales ($)')
plt.xlabel('Quarter')
plt.xticks(rotation=0)

# Chart 3: Seasonal sales by region (heatmap)
plt.subplot(2, 3, 3)
seasonal_region = sales_with_products.groupby(['Region', 'Quarter'])['Line Total'].sum().unstack(fill_value=0)
sns.heatmap(seasonal_region, annot=True, fmt='.0f', cmap='YlOrRd')
plt.title('Seasonal Sales by Region', fontweight='bold', fontsize=14)
plt.xlabel('Quarter')
plt.ylabel('Region')

# Chart 4: Monthly sales by region
plt.subplot(2, 3, 4)
monthly_region = sales_with_products.groupby(['Region', 'Month_Name'])['Line Total'].sum().unstack(fill_value=0)
monthly_region = monthly_region.reindex(columns=[
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
])
monthly_region.plot(kind='line', ax=plt.gca(), marker='o')
plt.title('Monthly Sales by Region', fontweight='bold', fontsize=14)
plt.ylabel('Total Sales ($)')
plt.xlabel('Month')
plt.xticks(rotation=45)
plt.legend(title='Region', bbox_to_anchor=(1.05, 1), loc='upper left')

# Chart 5: Seasonal profit analysis
plt.subplot(2, 3, 5)
quarterly_profit = sales_with_products.groupby('Quarter')['Profit'].sum()
quarterly_profit.plot(kind='bar', color='orange', alpha=0.7)
plt.title('Quarterly Profit', fontweight='bold', fontsize=14)
plt.ylabel('Total Profit ($)')
plt.xlabel('Quarter')
plt.xticks(rotation=0)

# Chart 6: Seasonal variance analysis
plt.subplot(2, 3, 6)
monthly_variance = monthly_sales / monthly_sales.mean() * 100
monthly_variance.plot(kind='bar', color='red', alpha=0.7)
plt.title('Monthly Sales Variance from Average (%)', fontweight='bold', fontsize=14)
plt.ylabel('Variance from Average (%)')
plt.xlabel('Month')
plt.xticks(rotation=45)
plt.axhline(y=100, color='black', linestyle='--', alpha=0.5, label='Average')
plt.legend()

plt.tight_layout()
plt.savefig('seasonal_analysis.png', dpi=300, bbox_inches='tight')
plt.show()

print("📊 Seasonal analysis saved as 'seasonal_analysis.png'")

# Find peak and low seasons
peak_month = monthly_sales.idxmax()
peak_sales = monthly_sales.max()
low_month = monthly_sales.idxmin()
low_sales = monthly_sales.min()

print(f"\n🎯 SEASONAL INSIGHTS:")
print(f"Peak month: {peak_month} with ${peak_sales:,.2f}")
print(f"Low month: {low_month} with ${low_sales:,.2f}")
print(f"Seasonal variance: {((peak_sales - low_sales) / low_sales * 100):.1f}%")

# Step 11: Seasonal Analysis for Specific Products
print("\n=== STEP 11: Seasonal Analysis for Specific Products ===")

# Analyze top 5 products seasonally
top_5_products = top_products.head(5).index
print(f"Analyzing seasonal patterns for top 5 products: {list(top_5_products)}")

# Create seasonal product analysis
plt.figure(figsize=(20, 15))

# Chart 1: Monthly sales for top 5 products
plt.subplot(2, 3, 1)
for product in top_5_products:
    product_monthly = sales_with_products[sales_with_products['Product Name'] == product].groupby('Month_Name')['Line Total'].sum()
    product_monthly = product_monthly.reindex([
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ])
    plt.plot(range(12), product_monthly.values, marker='o', label=product, linewidth=2)

plt.title('Monthly Sales for Top 5 Products', fontweight='bold', fontsize=14)
plt.ylabel('Sales ($)')
plt.xlabel('Month')
plt.xticks(range(12), ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], rotation=45)
plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
plt.grid(True, alpha=0.3)

# Chart 2: Quarterly sales for top 5 products
plt.subplot(2, 3, 2)
quarterly_products = sales_with_products[sales_with_products['Product Name'].isin(top_5_products)].groupby(['Product Name', 'Quarter'])['Line Total'].sum().unstack(fill_value=0)
quarterly_products.plot(kind='bar', ax=plt.gca())
plt.title('Quarterly Sales for Top 5 Products', fontweight='bold', fontsize=14)
plt.ylabel('Sales ($)')
plt.xlabel('Product')
plt.xticks(rotation=45)
plt.legend(title='Quarter')

# Chart 3: Seasonal variance by product
plt.subplot(2, 3, 3)
product_seasonal_variance = {}
for product in top_5_products:
    product_monthly = sales_with_products[sales_with_products['Product Name'] == product].groupby('Month_Name')['Line Total'].sum()
    product_monthly = product_monthly.reindex([
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ])
    variance = ((product_monthly.max() - product_monthly.min()) / product_monthly.min()) * 100
    product_seasonal_variance[product] = variance

variance_df = pd.DataFrame(list(product_seasonal_variance.items()), columns=['Product', 'Seasonal_Variance_%'])
sns.barplot(data=variance_df, x='Seasonal_Variance_%', y='Product', palette='viridis')
plt.title('Seasonal Variance by Product (%)', fontweight='bold', fontsize=14)
plt.xlabel('Seasonal Variance (%)')

# Chart 4: Product seasonal heatmap
plt.subplot(2, 3, 4)
product_seasonal = sales_with_products[sales_with_products['Product Name'].isin(top_5_products)].groupby(['Product Name', 'Quarter'])['Line Total'].sum().unstack(fill_value=0)
sns.heatmap(product_seasonal, annot=True, fmt='.0f', cmap='YlOrRd')
plt.title('Top 5 Products: Sales by Quarter', fontweight='bold', fontsize=14)
plt.xlabel('Quarter')
plt.ylabel('Product')

# Chart 5: Peak months by product
plt.subplot(2, 3, 5)
peak_months_by_product = {}
for product in top_5_products:
    product_monthly = sales_with_products[sales_with_products['Product Name'] == product].groupby('Month_Name')['Line Total'].sum()
    peak_month = product_monthly.idxmax()
    peak_months_by_product[product] = peak_month

peak_df = pd.DataFrame(list(peak_months_by_product.items()), columns=['Product', 'Peak_Month'])
peak_counts = peak_df['Peak_Month'].value_counts()
peak_counts.plot(kind='bar', color='orange', alpha=0.7)
plt.title('Peak Months Distribution', fontweight='bold', fontsize=14)
plt.ylabel('Number of Products')
plt.xlabel('Month')
plt.xticks(rotation=45)

# Chart 6: Product seasonality correlation
plt.subplot(2, 3, 6)
# Create correlation matrix for monthly sales between products
product_monthly_matrix = []
for product in top_5_products:
    product_monthly = sales_with_products[sales_with_products['Product Name'] == product].groupby('Month_Name')['Line Total'].sum()
    product_monthly = product_monthly.reindex([
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ])
    product_monthly_matrix.append(product_monthly.values)

correlation_matrix = pd.DataFrame(product_monthly_matrix, index=top_5_products).T.corr()
sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', center=0)
plt.title('Product Seasonality Correlation', fontweight='bold', fontsize=14)

plt.tight_layout()
plt.savefig('product_seasonal_analysis.png', dpi=300, bbox_inches='tight')
plt.show()

print("📊 Product seasonal analysis saved as 'product_seasonal_analysis.png'")

# Find products with different seasonal patterns
print(f"\n🎯 PRODUCT SEASONAL INSIGHTS:")
for product, variance in product_seasonal_variance.items():
    peak_month = peak_months_by_product[product]
    print(f"{product}: {variance:.1f}% seasonal variance, peak in {peak_month}")

# Check if any products have different seasonal patterns
unique_peak_months = set(peak_months_by_product.values())
if len(unique_peak_months) > 1:
    print(f"\n⚠️  DIFFERENT SEASONAL PATTERNS DETECTED!")
    print(f"Products peak in different months: {unique_peak_months}")
else:
    print(f"\n✅ CONSISTENT SEASONAL PATTERNS:")
    print(f"All top products peak in the same month: {list(unique_peak_months)[0]}")

# Step 12: Average Order Value per Household by Region
print("\n=== STEP 12: Average Order Value per Household by Region ===")

# Calculate average order value per household
print("Calculating average order value per household by region...")

# Group by region and calculate metrics
household_order_metrics = sales_with_products.groupby('Region').agg({
    'Line Total': 'mean',           # Average order value
    'OrderNumber': 'nunique',        # Number of unique orders
    'households': 'first'           # Number of households
}).reset_index()

# Calculate orders per household
household_order_metrics['Orders_per_Household'] = household_order_metrics['OrderNumber'] / household_order_metrics['households']
household_order_metrics['Avg_Order_Value'] = household_order_metrics['Line Total']
household_order_metrics['Total_Revenue_per_Household'] = (household_order_metrics['Orders_per_Household'] * household_order_metrics['Avg_Order_Value'])

print("\n📊 Order Value Analysis by Region:")
print(household_order_metrics[['Region', 'Avg_Order_Value', 'Orders_per_Household', 'Total_Revenue_per_Household']].round(2))

# Create visualization
plt.figure(figsize=(15, 10))

# Chart 1: Average Order Value by Region
plt.subplot(2, 3, 1)
sns.barplot(data=household_order_metrics, x='Region', y='Avg_Order_Value', palette='viridis')
plt.title('Average Order Value by Region', fontweight='bold', fontsize=14)
plt.ylabel('Average Order Value ($)')
plt.xticks(rotation=45)

# Chart 2: Orders per Household by Region
plt.subplot(2, 3, 2)
sns.barplot(data=household_order_metrics, x='Region', y='Orders_per_Household', palette='viridis')
plt.title('Orders per Household by Region', fontweight='bold', fontsize=14)
plt.ylabel('Orders per Household')
plt.xticks(rotation=45)

# Chart 3: Total Revenue per Household by Region
plt.subplot(2, 3, 3)
sns.barplot(data=household_order_metrics, x='Region', y='Total_Revenue_per_Household', palette='viridis')
plt.title('Total Revenue per Household by Region', fontweight='bold', fontsize=14)
plt.ylabel('Total Revenue per Household ($)')
plt.xticks(rotation=45)

# Chart 4: Scatter plot: Orders vs Order Value
plt.subplot(2, 3, 4)
plt.scatter(household_order_metrics['Orders_per_Household'], household_order_metrics['Avg_Order_Value'], 
           s=200, alpha=0.7, c=['red', 'blue', 'green', 'orange'])
for i, region in enumerate(household_order_metrics['Region']):
    plt.annotate(region, (household_order_metrics['Orders_per_Household'].iloc[i], 
                         household_order_metrics['Avg_Order_Value'].iloc[i]), 
                xytext=(5, 5), textcoords='offset points')
plt.title('Orders per Household vs Average Order Value', fontweight='bold', fontsize=14)
plt.xlabel('Orders per Household')
plt.ylabel('Average Order Value ($)')
plt.grid(True, alpha=0.3)

# Chart 5: Regional comparison (side by side)
plt.subplot(2, 3, 5)
x = range(len(household_order_metrics))
width = 0.25
plt.bar([i - width for i in x], household_order_metrics['Avg_Order_Value'], width, label='Avg Order Value', alpha=0.8)
plt.bar(x, household_order_metrics['Orders_per_Household'] * 1000, width, label='Orders per Household (×1000)', alpha=0.8)
plt.bar([i + width for i in x], household_order_metrics['Total_Revenue_per_Household'] / 1000, width, label='Revenue per Household (×1000)', alpha=0.8)
plt.title('Regional Metrics Comparison', fontweight='bold', fontsize=14)
plt.xlabel('Region')
plt.ylabel('Value')
plt.xticks(x, household_order_metrics['Region'], rotation=45)
plt.legend()

# Chart 6: Efficiency analysis
plt.subplot(2, 3, 6)
# Calculate efficiency score (revenue per household / orders per household)
household_order_metrics['Efficiency_Score'] = household_order_metrics['Total_Revenue_per_Household'] / household_order_metrics['Orders_per_Household']
sns.barplot(data=household_order_metrics, x='Region', y='Efficiency_Score', palette='viridis')
plt.title('Customer Efficiency Score by Region', fontweight='bold', fontsize=14)
plt.ylabel('Efficiency Score (Revenue per Order)')
plt.xticks(rotation=45)

plt.tight_layout()
plt.savefig('order_value_analysis.png', dpi=300, bbox_inches='tight')
plt.show()

