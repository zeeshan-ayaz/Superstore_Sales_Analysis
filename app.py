# Importing required libraries
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import openpyxl
import plotly.express as px
import streamlit as st

# Function to load default data
@st.cache_data
def load_default_data():
    return pd.read_excel("superstore.xlsx", engine='openpyxl')

# Sidebar for file upload or default dataset
st.sidebar.title("Upload or Load Dataset")

data_source = st.sidebar.radio(
    "Choose Data Source:",
    ("Default Dataset", "Upload Your Own Dataset")
)

# Load dataset based on user input
if data_source == "Default Dataset":
    df = load_default_data()
    st.sidebar.success("Default dataset loaded successfully!")
else:
    uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=['xlsx'])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        st.sidebar.success("Dataset uploaded successfully!")
    else:
        st.sidebar.warning("Please upload a dataset to proceed.")
        st.stop()  # Stop execution if no file is uploaded

# Convert order_date to datetime
df['order_date'] = pd.to_datetime(df['order_date'])

# Extract year and week from the order_date
df['year'] = df['order_date'].dt.year
df['week'] = df['order_date'].dt.isocalendar().week

# CSS for customization
st.markdown('<h1 class="title">Superstore Sales Analysis Report</h1>', unsafe_allow_html=True)

# CSS for customization
st.markdown("""
    <style>
    .stApp {
        background: #f4f4f4; /* Light grey background for the app */
        color: #333; /* Dark text color for better readability */
    }
    .css-1v0mbdj { 
        background-color: #ffffff !important; /* Sidebar background color */
        color: #333; /* Sidebar text color */
    }
    .stButton>button {
        background-color: #007bff; /* Primary button color */
        color: white; /* Button text color */
        width: 100%; /* Button width */
        border: none;
        border-radius: 8px;
        padding: 12px;
        font-size: 16px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #0056b3; /* Button hover color */
    }
    .title {
        color: #003366; /* Dark blue title color */
        font-size: 48px; /* Larger font size for titles */
        font-weight: bold; /* Bold title text */
        text-align: center; /* Center align the title */
        padding: 20px; /* Padding around the title */
        font-family: 'Arial', sans-serif; /* Font family for titles */
        text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.2); /* Title shadow */
    }
    h1 {
        color: #003366; /* Dark blue color for h1 */
    }
    h2 {
        color: #0056b3; /* Slightly lighter blue color for h2 */
    }
    h3 {
        color: #007bff; /* Lighter blue color for h3 */
    }
    .css-1d391kg { 
        width: 250px;  /* Adjust the sidebar width */
    }
    </style>
    """, unsafe_allow_html=True)
# Sidebar Navigation
st.sidebar.title("Navigation")

# Top level menu
page = st.sidebar.radio("Select a category", [
    "Overview",
    "Analysis",
    "Advanced Analysis"
])

if page == "Overview":
    subpage = st.sidebar.radio("Select Overview Page", [
        "Regional Performance Analysis",
        "Product Category Analysis"
    ])
elif page == "Analysis":
    subpage = st.sidebar.radio("Select Analysis Page", [
        "Customer Sales Analytics",
        "Profit Analytics",
        "Discount Strategy Analysis"
    ])
elif page == "Advanced Analysis":
    subpage = st.sidebar.radio("Select Advanced Analysis Page", [
        "Product Performance",
        "State-wise Performance",
        # "Sales Trends by Seasonality",
        "Correlation Analysis"
    ])

# Render the selected subpage
if subpage == "Regional Performance Analysis":
    st.header("Regional Performance Analysis")
    
    # First Plot: Total Sales by Region
    st.subheader("Total Sales by Region")
    
    # Aggregate total sales by region
    total_sales_by_region = df.groupby('region')['sales'].sum().reset_index()
    
    # Create a Plotly bar chart for total sales by region
    fig1 = px.bar(total_sales_by_region, 
                  x='region', 
                  y='sales', 
                  title='Total Sales by Region',
                  labels={'region': 'Region', 'sales': 'Total Sales'},
                  color='region',  # Color by region for better distinction
                  color_discrete_sequence=px.colors.qualitative.T10)
    
    # Customize layout
    fig1.update_layout(xaxis_title='Region', 
                       yaxis_title='Total Sales',
                       title_x=0.5, 
                       template='plotly_white',
                       width=700, 
                       height=500)
    
    # Display the plot in Streamlit
    st.plotly_chart(fig1)

    # Second Plot: Average Profit Margin by Region
    st.subheader("Average Profit Margin by Region")
    
    # Calculate average profit margin by region
    avg_profit_margin_by_region = df.groupby('region')['profit_margin'].mean().reset_index()
    
    # Create a Plotly bar chart for average profit margin by region
    fig2 = px.bar(avg_profit_margin_by_region, 
                  x='region', 
                  y='profit_margin', 
                  title='Average Profit Margin by Region',
                  labels={'region': 'Region', 'profit_margin': 'Average Profit Margin'},
                  color='region',  # Color by region for better distinction
                  color_discrete_sequence=px.colors.qualitative.T10)
    
    # Customize layout
    fig2.update_layout(xaxis_title='Region', 
                       yaxis_title='Average Profit Margin',
                       title_x=0.5, 
                       template='plotly_white',
                       width=700, 
                       height=500)
    
    # Display the plot in Streamlit
    st.plotly_chart(fig2)

elif subpage == "Product Category Analysis":
    st.header("Product Category Analysis")

    # Date filter at the top
    min_date = df['order_date'].min()
    max_date = df['order_date'].max()
    start_date = st.date_input("Start Date", value=min_date)
    end_date = st.date_input("End Date", value=max_date)

    # Ensure 'order_date' is in datetime format
    df['order_date'] = pd.to_datetime(df['order_date'])
    
    # Filter data based on the date range
    df_filtered = df[(df['order_date'] >= pd.to_datetime(start_date)) & (df['order_date'] <= pd.to_datetime(end_date))]

    # Region filter (optional)
    region_options = ["All Regions"] + list(df['region'].unique())
    selected_region = st.selectbox("Select Region", options=region_options, index=0)

    # Filter data based on the selected region (default to all regions if "All Regions" is selected)
    if selected_region != "All Regions":
        df_filtered = df_filtered[df_filtered['region'] == selected_region]

    # Category filter
    selected_category = st.selectbox("Select Category", options=df['category'].unique(), index=0)

    # Filter data based on the selected category
    category_sales = df_filtered[df_filtered['category'] == selected_category].groupby('subcategory')[['sales', 'profit']].sum().reset_index().sort_values(by='sales', ascending=False)

    # Display the sales by subcategory
    st.write(f"### Category: {selected_category} - Subcategories Sales and Profit (in {selected_region if selected_region != 'All Regions' else 'All Regions'})")
    st.dataframe(category_sales)

    # Plot sales and profit by subcategory
    fig2 = px.bar(category_sales, 
                  x='subcategory', 
                  y=['sales', 'profit'], 
                  title=f'Sales and Profit for {selected_category} in {selected_region if selected_region != "All Regions" else "All Regions"}', 
                  barmode='group',
                  labels={'sales': 'Total Sales', 'profit': 'Total Profit', 'subcategory': 'Subcategory'})

    # Update hover template for each trace
    fig2.for_each_trace(lambda t: t.update(hovertemplate=f'<b>{t.name}</b><br>Amount: %{t.y[0]:,.2f}<extra></extra>'))

    # Customize layout
    fig2.update_layout(xaxis_title="Subcategory", 
                       yaxis_title="Amount", 
                       title_x=0.5,
                       template='plotly_white')

    # Display the plot
    st.plotly_chart(fig2)

# Customer Sales Analytics Page
elif subpage == "Customer Sales Analytics":
    st.header("Customer Sales Analytics")

    # Show Total Number of Customers
    total_customers = df['customer'].nunique()
    st.subheader(f"Total Number of Customers: {total_customers}")

    # Display top 5 customers by profit
    st.subheader("Top 5 Customers by Profit")
    top_customers = df.groupby('customer')['profit'].sum().nlargest(5).reset_index()
    st.dataframe(top_customers)

    # Add date filter
    st.subheader("Filter by Date")
    start_date = st.date_input("Start Date", pd.to_datetime("2019-01-01"))
    end_date = st.date_input("End Date", pd.to_datetime("2020-12-31"))

    # Convert to datetime if necessary
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    # Filter dataset by date range
    mask = (df['order_date'] >= start_date) & (df['order_date'] <= end_date)
    filtered_data = df.loc[mask].copy()  # Ensure a copy of the filtered data

    if filtered_data.empty:
        st.warning("No data available for the selected date range.")
    else:
        # Select a customer to filter data
        selected_customer = st.selectbox("Select Customer", options=filtered_data['customer'].unique())
        customer_data = filtered_data.loc[filtered_data['customer'] == selected_customer].copy()

        st.subheader(f"Sales for Customer: {selected_customer}")

        # Display table for customer purchase details
        st.write("Purchase Details")
        st.dataframe(customer_data[['order_date', 'product_name', 'sales', 'quantity']])

        # Visualize sales by product for this customer
        product_sales = customer_data.groupby('product_name')['sales'].sum().reset_index()
        fig = px.bar(product_sales, y='product_name', x='sales', title=f'Sales by Product for {selected_customer}')
        st.plotly_chart(fig)

        # Visualize purchase history over time for this customer
        sales_over_time = customer_data.groupby('order_date')['sales'].sum().reset_index()
        fig = px.line(sales_over_time, x='order_date', y='sales', title=f'Sales Over Time for {selected_customer}', markers=True)
        st.plotly_chart(fig)


elif subpage == "Profit Analytics":
    st.header("Profit Analytics")
    
    # Display total profit across all regions by default
    total_profit = df['profit'].sum()
    st.write(f"**Total Profit Across All Regions:** ${total_profit:,.2f}")

    # Add date filter
    st.subheader("Select Date Range")
    start_date = st.date_input("From", df['order_date'].min().date())
    end_date = st.date_input("To", df['order_date'].max().date())

    # Filter the dataset based on the selected date range
    filtered_df = df[(df['order_date'] >= pd.to_datetime(start_date)) & (df['order_date'] <= pd.to_datetime(end_date))]

    # Filter for region selection
    selected_profit_region = st.selectbox("Select Region", options=["All"] + list(filtered_df['region'].unique()))
    
    if selected_profit_region == "All":
        region_profit = filtered_df.groupby('region')['profit'].sum().reset_index()
        st.write(f"**Total Profit for Selected Date Range:** ${region_profit['profit'].sum():,.2f}")
    else:
        region_profit = filtered_df[filtered_df['region'] == selected_profit_region].groupby('region')['profit'].sum().reset_index()
        st.write(f"**Total Profit for {selected_profit_region} in Selected Date Range:** ${region_profit['profit'].values[0]:,.2f}")
    
    fig = px.bar(region_profit, x='region', y='profit', title=f'Total Profit in {selected_profit_region if selected_profit_region != "All" else "All Regions"}', color='region')
    st.plotly_chart(fig)
    
    # Add filter for category and subcategory
    selected_category = st.selectbox("Select Category", options=["All"] + list(filtered_df['category'].unique()))
    
    if selected_category == "All":
        filtered_data = filtered_df
    else:
        filtered_data = filtered_df[filtered_df['category'] == selected_category]
    
    st.subheader(f"Profit by Sub-Category in {selected_profit_region if selected_profit_region != 'All' else 'All Regions'} for {selected_category if selected_category != 'All' else 'All Categories'}")
    subcategory_profit = filtered_data.groupby('subcategory')['profit'].sum().reset_index()
    fig = px.bar(subcategory_profit, x='subcategory', y='profit', title=f'Profit by Sub-Category in {selected_profit_region if selected_profit_region != "All" else "All Regions"}', color='subcategory')
    st.plotly_chart(fig)


# Discount Analytics
elif subpage == "Discount Strategy Analysis":
    st.header("Discount Strategy Analysis")
    st.subheader("Discount Impact on Sales and Profit")

    # Extract Year and Month from 'order_date' column
    df['order_date'] = pd.to_datetime(df['order_date'])  # Ensure 'order_date' is in datetime format

    # Date filter
    st.write("**Select Date Range**")
    start_date = st.date_input("From", min_value=df['order_date'].min().date(), max_value=df['order_date'].max().date())
    end_date = st.date_input("To", min_value=start_date, max_value=df['order_date'].max().date())

    # Region filter
    region_options = [None] + list(df['region'].unique())
    selected_region = st.selectbox("Select Region (Optional)", options=region_options)

    # Filter data based on the selected date range and region
    filtered_df = df[(df['order_date'] >= pd.to_datetime(start_date)) & (df['order_date'] <= pd.to_datetime(end_date))]
    if selected_region:
        filtered_df = filtered_df[filtered_df['region'] == selected_region]

    # Show overall discount impact (if no filter is applied)
    st.write("### Overall Discount Strategy Impact on Sales and Profit")
    overall_discount_impact = df.groupby('discount')[['sales', 'profit']].sum().reset_index()

    # Show overall discount impact using a line chart
    fig_overall = px.line(overall_discount_impact, x='discount', y=['sales', 'profit'],
                          title="Overall Sales and Profit by Discount",
                          labels={'sales': 'Total Sales', 'profit': 'Total Profit'},
                          markers=True)
    fig_overall.update_traces(mode='lines+markers')
    fig_overall.update_layout(
        xaxis_title='Discount',
        yaxis_title='Amount',
        legend_title='Metrics'
    )
    # Customize colors for the lines
    fig_overall.update_traces(line=dict(color='blue'), selector=dict(name='sales'))
    fig_overall.update_traces(line=dict(color='red'), selector=dict(name='profit'))
    
    # Add hover data to display detailed information
    fig_overall.update_traces(
        hovertemplate='Discount: %{x}<br>Sales: %{y}<br>Profit: %{customdata[1]}<extra></extra>',
        customdata=overall_discount_impact[['discount', 'profit']].values
    )
    st.plotly_chart(fig_overall)

    
# Product Performance Analysis
elif subpage == "Product Performance":
    st.header("Product Performance")

    # Add date filter and region filter at the top
    start_date = st.date_input("Start Date", value=pd.to_datetime("2019-01-01"))
    end_date = st.date_input("End Date", value=pd.to_datetime("2020-12-31"))
    
    # Ensure the 'order_date' column is in datetime format
    df['order_date'] = pd.to_datetime(df['order_date'])
    
    # Filter data by date range
    df_filtered = df[(df['order_date'] >= pd.to_datetime(start_date)) & (df['order_date'] <= pd.to_datetime(end_date))]

    # Region filter (optional)
    region_options = ["All Regions"] + list(df['region'].unique())
    selected_region = st.selectbox("Select Region", options=region_options, index=0)

    # Filter data based on the selected region (if region is selected)
    if selected_region != "All Regions":
        filtered_df = df_filtered[df_filtered['region'] == selected_region]
    else:
        filtered_df = df_filtered  # If no region is selected, use the full dataset

    # Add button to toggle between top 5 highest and lowest products
    toggle_button = st.radio("Choose Products to Display", ("Top 5 Performing Products", "Top 5 Lowest Performing Products"))

    if toggle_button == "Top 5 Performing Products":
        # Top 5 Products by Sales
        st.subheader("Top 5 Products by Sales")
        top_products = filtered_df.groupby('product_name')['sales'].sum().reset_index().sort_values(by='sales', ascending=False).head(5)

        # Show the top products in a table
        st.dataframe(top_products)

        # Visualize the top 5 products by sales using a bar chart
        fig_top_products = px.bar(top_products, y='product_name', x='sales', title=f"Top 5 Products by Sales in {selected_region if selected_region != 'All Regions' else 'All Regions'}",
                                  labels={'sales': 'Total Sales', 'product_name': 'Product'})
        st.plotly_chart(fig_top_products)

        # Profit for Top 5 Products
        st.subheader("Profit for Top 5 Products")
        top_products_profit = filtered_df[filtered_df['product_name'].isin(top_products['product_name'])]
        top_products_profit = top_products_profit.groupby('product_name')['profit'].sum().reset_index()

        # Show the profit for the top products in a table
        st.dataframe(top_products_profit)

        # Visualize the profit for top 5 products using a bar chart
        fig_top_profit = px.bar(top_products_profit, y='product_name', x='profit', title=f"Profit for Top 5 Products in {selected_region if selected_region != 'All Regions' else 'All Regions'}",
                                labels={'profit': 'Total Profit', 'product_name': 'Product'}, )
        st.plotly_chart(fig_top_profit)

    elif toggle_button == "Top 5 Lowest Performing Products":
        # Top 5 Lowest Products by Sales (loss-making products)
        st.subheader("Top 5 Lowest Performing Products by Sales")
        bottom_products = filtered_df.groupby('product_name')['sales'].sum().reset_index().sort_values(by='sales', ascending=True).head(5)

        # Show the lowest performing products in a table
        st.dataframe(bottom_products)

        # Visualize the bottom 5 products by sales using a bar chart
        fig_bottom_products = px.bar(bottom_products, y='product_name', x='sales', title=f"Bottom 5 Products by Sales in {selected_region if selected_region != 'All Regions' else 'All Regions'}",
                                     labels={'sales': 'Total Sales', 'product_name': 'Product'})
        st.plotly_chart(fig_bottom_products)

        # Profit for Bottom 5 Products
        st.subheader("Profit/Loss for Bottom 5 Products")
        bottom_products_profit = filtered_df[filtered_df['product_name'].isin(bottom_products['product_name'])]
        bottom_products_profit = bottom_products_profit.groupby('product_name')['profit'].sum().reset_index()

        # Show the profit/loss for the bottom products in a table
        st.dataframe(bottom_products_profit)

        # Visualize the profit/loss for bottom 5 products using a bar chart
        fig_bottom_profit = px.bar(bottom_products_profit, y='product_name', x='profit', title=f"Profit/Loss for Bottom 5 Products in {selected_region if selected_region != 'All Regions' else 'All Regions'}",
                                   labels={'profit': 'Profit/Loss', 'product_name': 'Product'}, )
        st.plotly_chart(fig_bottom_profit)


# State-wise Performance Analysis
elif subpage == "State-wise Performance":
    st.header("State-wise Performance")

    # Add date filter and region filter at the top
    min_date = df['order_date'].min()
    max_date = df['order_date'].max()

    start_date = st.date_input("Start Date", value=min_date)
    end_date = st.date_input("End Date", value=max_date)

    # Ensure 'order_date' is in datetime format
    df['order_date'] = pd.to_datetime(df['order_date'])

    # Filter data by date range (default to all time if no specific range is selected)
    df_filtered = df[(df['order_date'] >= pd.to_datetime(start_date)) & (df['order_date'] <= pd.to_datetime(end_date))]

    # Region filter (optional) with "All Regions" as default
    region_options = ["All Regions"] + list(df['region'].unique())
    selected_region = st.selectbox("Select Region", options=region_options, index=0)

    # Filter data based on the selected region (default to all regions if "All Regions" is selected)
    if selected_region != "All Regions":
        df_filtered = df_filtered[df_filtered['region'] == selected_region]

    # Total Quantity Sold by State
    st.subheader("Total Quantity Sold by State")

    # Group data by state and calculate total quantity sold
    product_sales = df_filtered.groupby('state')['quantity'].sum().reset_index()

    # Create a horizontal bar chart with Plotly
    fig = px.bar(product_sales, 
                 x='quantity', 
                 y='state', 
                 title=f'Total Quantity Sold by State in {selected_region if selected_region != "All Regions" else "All Regions"}',
                 labels={'quantity': 'Total Quantity', 'state': 'State'},
                 orientation='h',  # Horizontal bar chart
                 color='quantity',  # Color by quantity for better visualization
                 color_continuous_scale='viridis')  # Use the 'viridis' color palette

    # Customize layout
    fig.update_layout(xaxis_title='Total Quantity', 
                      yaxis_title='State',
                      title_x=0.5, 
                      template='plotly_white',
                      width=800, 
                      height=600)

    # Display the plot in Streamlit
    st.plotly_chart(fig)


# # Sales Trends by Seasonality
# elif subpage == "Sales Trends by Seasonality":
#     st.header("Sales Trends by Seasonality")
    
#     # Monthly Sales Trend
#     st.subheader("Monthly Sales Trend")
#     df['order_date'] = pd.to_datetime(df['order_date'])
#     df['month'] = df['order_date'].dt.month
#     monthly_sales = df.groupby('month')['sales'].sum().reset_index()
#     fig = px.line(monthly_sales, x='month', y='sales', title='Monthly Sales Trend')
#     st.plotly_chart(fig)

# Correlation Analysis
elif subpage == "Correlation Analysis":
    st.header("Correlation Analysis")
    
    # Display correlation heatmap using seaborn
    st.subheader("Correlation Heatmap")
    corr = df[['sales', 'profit', 'quantity', 'discount']].corr()
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.heatmap(corr, annot=True, cmap='coolwarm', linewidths=0.5, ax=ax)
    ax.set_title('Correlation Matrix')
    st.pyplot(fig) 
