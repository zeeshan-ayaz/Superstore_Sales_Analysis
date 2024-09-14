import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import openpyxl 
import plotly.express as px
import streamlit as st

# Function to load data
@st.cache_data
def load_data():
    return pd.read_excel("superstore.xlsx", engine='openpyxl')

# Load data
df = load_data()

st.markdown('<h1 class="title">Superstore Sales Analysis Report</h1>', unsafe_allow_html=True)


# Set the background color, sidebar width, and button styling using custom CSS
st.markdown("""
    <style>
    .stApp {
        background-color: #f0f2f6;
    }
    .css-1d391kg { 
        width: 200px;  /* Adjust the width here */
    }
    .stButton>button {
        background-color: #4CAF50; /* Button background color */
        color: white; /* Button text color */
        width: 100%; /* Button width */
        border: none;
        border-radius: 8px;
        padding: 16px;
        font-size: 25px;
        font-weight:bold;
    }
    .stButton>button:hover {
        background-color: #45a049; /* Button hover color */
    }
    .title {
        color: #FF0000; /* Title color red */
        font-size: 45px; /* Title font size */
        font-weight: bold; /* Title font weight */
        text-align: center; /* Center align the title */
        padding: 20px; /* Add padding around the title */
        font-family: 'Arial', sans-serif; /* Title font family */
        text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.2); /* Add shadow for a 3D effect */
    }
    .css-1v0mbdj { /* Class for sidebar */
        background-color: #FFFF00; /* Sidebar background color yellow */
    }
    </style>
    """, unsafe_allow_html=True)

# Button to reload data
if st.sidebar.button('Reload Data'):
    df = load_data()
    st.sidebar.success('Data reloaded successfully!')

# Sidebar
st.sidebar.title("Navigation")
page = st.sidebar.radio("Select a page", [
    "Regional Performance Analysis",
    "Product Category Analysis",
    "Sales Trends by Seasonality",
    "Discount Strategy Analysis",
    "Product Performance",
    "State-wise Performance",
    "Correlation Analysis"
])


# Regional Performance Analysis
if page == "Regional Performance Analysis":
    st.header("1. Regional Performance Analysis")
    
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


# Product Category Analysis
elif page == "Product Category Analysis":
    st.header("2. Product Category Analysis")
    
    # First Plot: Profit Margin by Category
    st.subheader("Profit Margin by Category")
    
    # Calculate average profit margin by category
    profit_margin_by_category = df.groupby('category')['profit_margin'].mean().reset_index()
    
    # Create a Plotly bar chart for profit margins
    fig1 = px.bar(profit_margin_by_category, 
                  x='category', 
                  y='profit_margin', 
                  title='Profit Margin by Category',
                  labels={'category': 'Category', 'profit_margin': 'Average Profit Margin'},
                  color='category',  # Add color by category for better distinction
                  color_discrete_sequence=px.colors.qualitative.T10)
    
    # Customize layout
    fig1.update_layout(xaxis_title='Category', 
                       yaxis_title='Average Profit Margin',
                       title_x=0.5, 
                       template='plotly_white',
                       width=700, 
                       height=500)
    
    # Display the plot in Streamlit
    st.plotly_chart(fig1)

    # Second Plot: Sales by Category and Sub-Category
    st.subheader("Sales by Category and Sub-Category")
    
    # Group data by category and subcategory to calculate total sales
    grouped_data = df.groupby(['category', 'subcategory'])['sales'].sum().reset_index()
    
    # Create a Plotly stacked bar chart
    fig2 = px.bar(grouped_data, 
                  x='category', 
                  y='sales', 
                  color='subcategory', 
                  title='Sales by Category and Sub-Category',
                  labels={'category': 'Category', 'sales': 'Sales'},
                  color_discrete_sequence=px.colors.qualitative.T10,
                  barmode='stack')  # Stacked bars for sub-category
    
    # Customize layout
    fig2.update_layout(xaxis_title='Category', 
                       yaxis_title='Sales',
                       title_x=0.5, 
                       xaxis_tickangle=-45,  # Rotate x-axis labels for better readability
                       legend_title_text='Sub-Category',
                       legend=dict(x=1.05, y=1),  # Position legend outside the plot
                       template='plotly_white',
                       width=700, 
                       height=500)
    
    # Display the plot in Streamlit
    st.plotly_chart(fig2)


# Sales Trends by Seasonality
# Assuming you have the df dataframe ready
elif page == "Sales Trends by Seasonality":
    st.header("3. Sales Trends by Seasonality")
    
    st.subheader("Monthly Sales Over Time")
    
    # Convert order_date to datetime
    df['order_date'] = pd.to_datetime(df['order_date'])
    
    # Extract year and month separately
    df['year'] = df['order_date'].dt.year
    df['year_month'] = df['order_date'].dt.to_period('M')
    
    # Get unique years from the dataset
    years = df['year'].unique()
    
    # Create a dropdown or radio button for selecting a year
    selected_year = st.selectbox('Select Year', sorted(years))
    
    # Filter the dataframe for the selected year
    df_filtered = df[df['year'] == selected_year]
    
    # Aggregate sales by year-month for the selected year
    monthly_sales = df_filtered.groupby('year_month')['sales'].sum().reset_index()
    
    # Convert year_month to string for Plotly
    monthly_sales['year_month'] = monthly_sales['year_month'].astype(str)
    
    # Plot with Plotly Express
    fig = px.line(monthly_sales, 
                  x='year_month', 
                  y='sales', 
                  title=f'Monthly Sales Over Time for {selected_year}',
                  markers=True,
                  labels={'year_month': 'Date (Year-Month)', 'sales': 'Total Sales'},
                  line_shape='linear',
                  color_discrete_sequence=['red'])  # Add a color for consistency
    
    # Customize layout
    fig.update_layout(
        xaxis_title=f'Year {selected_year}',
        yaxis_title='Total Sales',
        title_x=0.5,
        xaxis_tickangle=-45,
        template='plotly_white',
        width=900, 
        height=500
    )
    
    # Display the plot in Streamlit
    st.plotly_chart(fig)


# Discount Strategy Analysis
elif page == "Discount Strategy Analysis":
    st.header("4. Discount Strategy Analysis")
    
    # Plot: Distribution of Discounts
    st.subheader("Distribution of Discounts")
    
    # Create a Plotly histogram for discount distribution
    fig = px.histogram(df, 
                       x='discount', 
                       nbins=30,  # size of bins
                       title='Distribution of Discounts',
                       labels={'discount': 'Discount', 'count': 'Frequency'},
                       color_discrete_sequence=['indianred'])  # Choose color for better distinction
    
    # Customize layout
    fig.update_layout(xaxis_title='Discount', 
                      yaxis_title='Frequency',
                      title_x=0.5, 
                      template='plotly_white',
                      width=700, 
                      height=500)
    
    # Display the plot in Streamlit
    st.plotly_chart(fig)


# Product Performance
elif page == "Product Performance":
    st.header("5. Product Performance")
    
    st.subheader("Top 10 Products by Sales")
    product_sales = df.groupby('product_name')['sales'].sum().reset_index()
    top_10_products = product_sales.sort_values(by='sales', ascending=False).head(10)
    fig, ax = plt.subplots(figsize=(10, 8))
    sns.barplot(y='product_name', x='sales', data=top_10_products, palette='viridis', ax=ax)
    ax.set_title('Top 10 Products by Sales')
    ax.set_xlabel('Total Sales')
    ax.set_ylabel('Product Name')
    st.pyplot(fig)


# State-wise Performance
elif page == "State-wise Performance":
    st.header("6. State-wise Performance")
    
    # Total Quantity Sold by State
    st.subheader("Total Quantity Sold by State")
    
    # Group data by state and calculate total quantity sold
    product_sales = df.groupby('state')['quantity'].sum().reset_index()
    
    # Create a horizontal bar chart with Plotly
    fig = px.bar(product_sales, 
                 x='quantity', 
                 y='state', 
                 title='Total Quantity Sold by State',
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


# Correlation Analysis
elif page == "Correlation Analysis":
    st.header("7. Correlation Analysis")
    
    st.subheader("Correlation Heatmap")
    numeric_cols = ['profit_margin', 'sales', 'profit', 'discount', 'quantity'] 
    corr_matrix = df[numeric_cols].corr()
    fig, ax = plt.subplots(figsize=(8, 6))
    sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', linewidths=0.5, fmt=".2f", ax=ax)
    ax.set_title('Correlation Heatmap')
    st.pyplot(fig)
