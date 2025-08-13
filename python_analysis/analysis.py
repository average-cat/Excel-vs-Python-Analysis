def main():

    import numpy as np
    import pandas as pd
    import matplotlib.pyplot as plt
    import seaborn as sns
    import openpyxl
    from pathlib import Path
    import pandas as pd

    BASE = Path(__file__).resolve().parent.parent
    data_path = BASE / "data" / "USSuperstoredata.xlsx"
    df = pd.read_excel(data_path, engine="openpyxl")
    
    df.head()

    # MONTHLY SALES TREND

    df['Order Date'] = pd.to_datetime(df['Order Date'])

    df['Month Year'] = df['Order Date'].dt.to_period('M').astype(str)
    df.head()

    monthly_sales = df.groupby('Month Year')['Sales'].sum().reset_index()

    plt.figure(figsize=(12,6))
    plt.plot(monthly_sales['Month Year'], monthly_sales['Sales'], marker='o')
    plt.xticks(rotation=45)
    plt.title("Monthly Sales Trend")
    plt.xlabel("Month-Year")
    plt.ylabel("Sales")
    plt.tight_layout()
    plt.show()

    # TOP SUB-CATEGORIES BY SALE

    subcat_sales = df.groupby('Sub-Category')['Sales'].sum().sort_values(ascending=False)
    subcat_sales.plot(kind='bar', figsize=(10,5), title="Sub-Category Sales")
    plt.ylabel("Sales")
    plt.tight_layout()
    plt.show()

    # SEGMENT-WISE SALES

    seg_sales = df.groupby('Segment')['Sales'].sum()

    seg_sales.plot(kind='pie', autopct='%1.1f%%', startangle=90, title="Sales by Segment", colors = ['salmon', 'limegreen', 'cornflowerblue'])
    plt.ylabel("")
    plt.show()

    # REGION-WISE SALES

    reg_sales = df.groupby("Region")['Sales'].sum()

    reg_sales.plot(kind='bar', title="Sales by Region", color='cornflowerblue', alpha=0.8)
    plt.ylabel("Sales")
    plt.show()

    # PROFIT ANALYSIS

    check_profit = [feature for feature in df.columns if feature in ['Segment', 'Region', 'Category', 'Month Year']]
    for feature in check_profit:
        feature_profit = df.groupby(feature)['Profit'].sum().sort_values(ascending=False)
        feature_profit.plot(kind='bar', title = f'Sales by {feature}', figsize=(10,6), color='hotpink', alpha=0.9)
        plt.ylabel("Profit")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

    # PROFIT-TO-SALES RATIO BY CATEGORY

    category_profit_sales = df.groupby('Category')[['Sales', 'Profit']].sum()
    category_profit_sales['Profit-to-Sales%'] = 100 * category_profit_sales['Profit']/category_profit_sales['Sales']
    category_profit_sales = category_profit_sales.sort_values('Profit-to-Sales%', ascending=False)

    category_profit_sales['Profit-to-Sales%'].plot(kind='barh', title="Profit-to-Sales Ratio by Category", color='turquoise')

    # YEAR-O-YEAR (YoY) GROWTH

    df['Year'] = df['Order Date'].dt.year
    yoy = df.groupby('Year')[['Sales', 'Profit']].sum()

    yoy['Sales Growth %'] = yoy['Sales'].pct_change() * 100
    yoy['Profit Growth %'] = yoy['Profit'].pct_change() * 100

    yoy[['Sales Growth %', 'Profit Growth %']].plot(kind='bar', title='YoY Growth (%)')

if __name__ == "__main__":
    main()