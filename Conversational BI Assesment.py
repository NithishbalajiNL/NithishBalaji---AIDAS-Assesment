import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import streamlit as st

@st.cache_data
def load_data():
    try:
        df = pd.read_excel("Sales_data (1).xlsx")

        df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
        df["Order Release Date"] = pd.to_datetime(df["Order Release Date"], errors="coerce")
        df["Date Shipped"] = pd.to_datetime(df["Date Shipped"], errors="coerce")

        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()

df = load_data()

if "messages" not in st.session_state or not isinstance(st.session_state.messages, list):
    st.session_state.messages = []

def top_selling_items(df):
    return df.groupby("Item ID")["Total Price"].sum().nlargest(5)

def customer_with_most_returns(df):
    return df.groupby("Customer ID")["Qty Returned"].sum().idxmax()

def monthly_sales_trend(df):
    df["Month"] = df["Order Date"].dt.to_period("M")
    sales_trend = df.groupby("Month")["Total Price"].sum().reset_index()

    fig, ax = plt.subplots(figsize=(12, 6))
    sns.lineplot(x=sales_trend["Month"].astype(str), y=sales_trend["Total Price"], marker="o", ax=ax)
    plt.xticks(rotation=45)
    plt.title("📈 Monthly Sales Trend")
    return fig

def total_sales(df):
    return df["Total Price"].sum()

def average_order_value(df):
    return df["Total Price"].mean()

def most_frequent_customer(df):
    return df["Customer ID"].mode()[0]

def least_selling_items(df):
    return df.groupby("Item ID")["Total Price"].sum().nsmallest(5)

def top_customers_by_spending(df):
    return df.groupby("Customer ID")["Total Price"].sum().nlargest(5)

def products_with_highest_quantity_sold(df):
    return df.groupby("Item ID")["Qty Ordered"].sum().nlargest(5)

def busiest_sales_day(df):
    return df["Order Date"].value_counts().idxmax()

def most_returned_item(df):
    return df.groupby("Item ID")["Qty Returned"].sum().idxmax()

def yearly_sales_trend(df):
    df["Year"] = df["Order Date"].dt.year
    sales_trend = df.groupby("Year")["Total Price"].sum().reset_index()

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(x=sales_trend["Year"], y=sales_trend["Total Price"], ax=ax)
    plt.title("Yearly Sales Trend")
    return fig

def most_ordered(df):
    return df.groupby("Item ID")["Qty Ordered"].sum().idxmax()

def orders_shipped_this_week(df):
    today = pd.to_datetime("today")
    start_of_week = today - pd.DateOffset(days=today.weekday())
    end_of_week = start_of_week + pd.DateOffset(days=6)

    shipped_this_week = df[(df["Date Shipped"] >= start_of_week) & (df["Date Shipped"] <= end_of_week)]
    return len(shipped_this_week)

def top_customers_by_revenue(df, top_n=5):
    return df.groupby("Customer ID")["Total Price"].sum().nlargest(top_n)

def most_used_shipping_method(df):
    if "Ship Code" in df.columns:
        return df["Ship Code"].mode()[0]
    return "Shipping method data not available."

def bottom_customers_by_revenue(df, n=5):
    if "Customer ID" in df.columns and "Total Price" in df.columns:
        return df.groupby("Customer ID")["Total Price"].sum().nsmallest(n)
    return "Customer or revenue data not available."
def customer_return_chart(df):
    customer_returns = df.groupby("Customer ID")["Qty Returned"].sum().nlargest(5).reset_index()
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.barplot(x="Customer ID", y="Qty Returned", data=customer_returns, ax=ax)
    ax.set_title("Top 5 Customers with Most Returns")
    return fig


def shipping_method_pie_chart(df):
    shipping_counts = df["Ship Code"].value_counts()
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.pie(shipping_counts, labels=shipping_counts.index, autopct="%1.1f%%", startangle=90)
    ax.set_title("Most Used Shipping Methods")
    return fig

def top_customers_chart(df):
    top_customers = df.groupby("Customer ID")["Total Price"].sum().nlargest(5).reset_index()
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.barplot(x="Customer ID", y=" Total Price", data=top_customers, ax=ax)
    ax.set_title("Top 5 Revenue-Contributing Customers")
    return fig

def bottom_customers_chart(df):
    bottom_customers = df.groupby("Customer ID")["Total Price"].sum().nsmallest(5).reset_index()
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.barplot(x="Customer ID", y="Total Price", data=bottom_customers, ax=ax)
    ax.set_title("Bottom 5 Revenue-Contributing Customers")
    return fig
def top_selling_products_chart(df):
    top_selling = df.groupby("Item ID")["Total Price"].sum().nlargest(5).reset_index()
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.barplot(x="Item ID", y="Total Price", data=top_selling, ax=ax)
    ax.set_title("Top 5 Selling Products")
    return fig

st.title("Business Intelligence Chatbot")
st.write("Hey there! I'm your BI Assistant. Ask me anything about your sales, customers, and trends!")

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

query = st.chat_input("Ask me a business question...")

if query:
    st.session_state.messages.append({"role": "user", "content": query})
    query_lower = query.lower()
    response = ""

    if "what are the top 5 selling items" in query_lower:
        response = top_selling_items(df).to_frame().to_markdown()
    elif "show me customer with most returns" in query_lower:
        response = f"**Customer with Most Returns:** `{customer_with_most_returns(df)}`"
    elif "show me monthly sales trends" in query_lower:
        response = "**Here is the Monthly Sales Trend:**"
        st.pyplot(monthly_sales_trend(df))
    elif "show me total sales" in query_lower:
        response = f"**Total Sales:** `${total_sales(df):,.2f}`"
    elif "what is the average order value" in query_lower:
        response = f"**Average Order Value:** `${average_order_value(df):,.2f}`"
    elif "who is the most frequent customer" in query_lower:
        response = f"**Most Frequent Customer:** `{most_frequent_customer(df)}`"
    elif "show me the least selling items" in query_lower:
        response = least_selling_items(df).to_frame().to_markdown()
    elif "who are top customers by spending" in query_lower:
        response = top_customers_by_spending(df).to_frame().to_markdown()
    elif "show me products with highest quantity sold" in query_lower:
        response = products_with_highest_quantity_sold(df).to_frame().to_markdown()
    elif "show me busiest sales day" in query_lower:
        response = f"**Busiest Sales Day:** `{busiest_sales_day(df)}`"
    elif "show me the most returned item" in query_lower:
        response = f"**Most Returned Item:** `{most_returned_item(df)}`"
    elif "generate a graph to show yearly sales trends" in query_lower:
        response = "**Here is the Yearly Sales Trend:**"
        st.pyplot(yearly_sales_trend(df))
    elif "show me most ordered product" in query_lower:
        response = f"Most ordered item is: `{most_ordered(df)}`"
    elif "show me the orders shipped this week" in query_lower:
        response = f"Orders shipped this week: `{orders_shipped_this_week(df)}`"
    elif "show top customers contributing to revenue" in query_lower:
        response = top_customers_by_revenue(df).to_frame().to_markdown()
    elif "which shipping method was used the most" in query_lower:
        response = f"Most used Shipping method: `{most_used_shipping_method(df)}`"
    elif "show bottom customers contributing least to revenue" in query_lower:
        response = bottom_customers_by_revenue(df).to_frame().to_markdown()
    elif "generate a graph to show customer with most returns" in query_lower:
        response = "Here are the top customers with most returns:"
        st.pyplot(customer_return_chart(df))

    elif "generate a graph to show monthly sales trends" in query_lower:
        response = "Here is the Monthly Sales Trend:"
        st.pyplot(monthly_sales_trend(df))

    elif "generate a graph to show yearly sales trends" in query_lower:
        response = "Here is the Yearly Sales Trend:"
        st.pyplot(yearly_sales_trend(df))

    elif "generate a graph to show top customers contributing to revenue" in query_lower:
        response = "Here are the top customers contributing to revenue:"
        st.pyplot(top_customers_chart(df))

    elif "generate a graph to show bottom customers contributing the least to revenue" in query_lower:
        response = "Here are the bottom customers contributing the least to revenue:"
        st.pyplot(bottom_customers_chart(df))

    elif "generate a graph to show which shipping method was used the most" in query_lower:
        response = "Here is the most used shipping method distribution:"
        st.pyplot(shipping_method_pie_chart(df))
    elif  "generate a graph to show top 5 selling items" in query_lower:
        response = "Here are the top 5 selling products:"
        st.pyplot(top_selling_products_chart(df))
    else:
        response = "Sorry, I didn't understand your question. Try asking about sales trends, top customers, or order details."

    
    st.session_state.messages.append({"role": "assistant", "content": response})
    with st.chat_message("assistant"):
        st.markdown(response)
