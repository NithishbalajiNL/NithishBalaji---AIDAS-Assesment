import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import streamlit as st
import spacy

nlp = spacy.load("en_core_web_sm")
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
def extract_keywords(text):
    doc = nlp(text)
    keywords = [token.lemma_.lower() for token in doc if token.is_alpha and not token.is_stop]
    return " ".join(keywords)


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
    df["Year"] = df["Order Date"].dt.year.astype("Int64")  # Ensuring integer values
    sales_trend = df.groupby("Year")["Total Price"].sum().reset_index()

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.lineplot(x=sales_trend["Year"], y=sales_trend["Total Price"], ax=ax, marker="o")
    plt.xticks(sales_trend["Year"])  # Ensure proper year labeling
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
    sns.barplot(x="Customer ID", y="Total Price", data=top_customers, ax=ax)
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
def top_customers_by_spending_chart(df):
    top_customers = df.groupby("Customer ID")["Total Price"].sum().nlargest(5).reset_index()
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.barplot(x="Customer ID", y="Total Price", data=top_customers, ax=ax)
    ax.set_title("Top 5 Customers by Spending")
    return fig

def classify_query(query):
    query = query.lower()
    if any(word in query for word in ["chart", "graph", "plot", "trend"]):
        return "chart"
    return "table"


# Table Data Functions
def get_table_data(query):
    query = query.lower()  # Normalize query for case-insensitive matching
    if "top selling items" in query:
        return top_selling_items(df)
    elif "customer with most returns" in query:
        return customer_with_most_returns(df)
    elif "total sales" in query:
        return total_sales(df)
    elif "average order value" in query:
        return average_order_value(df)
    elif "most frequent customer" in query:
        return most_frequent_customer(df)
    elif "least selling items" in query:
        return least_selling_items(df)
    elif "top customers by spending" in query:
        return top_customers_by_spending(df)
    elif "products with highest quantity sold" in query:
        return products_with_highest_quantity_sold(df)
    elif "busiest sales day" in query:
        return busiest_sales_day(df)
    elif "most returned item" in query:
        return most_returned_item(df)
    elif "most ordered" in query:
        return most_ordered(df)
    elif "orders shipped this week" in query:
        return orders_shipped_this_week(df)
    elif "top customers contributing to revenue" in query:
        return top_customers_by_revenue(df)
    elif "shipping method used the most" in query:
        return most_used_shipping_method(df)
    elif "bottom customers by revenue" in query:
        return bottom_customers_by_revenue(df)
    else:
        return "Sorry, I couldn't find the information you are looking for."


# Chart Functions
def get_chart(query):
    if "monthly sales trend" in query:
        return monthly_sales_trend(df)
    elif "yearly sales trend" in query:
        return yearly_sales_trend(df)
    elif "customer with most returns" in query:
        return customer_return_chart(df)
    elif "top customers contributing to revenue" in query:
        return top_customers_chart(df)
    elif "bottom customers contributing to revenue" in query:
        return bottom_customers_chart(df)
    elif "shipping method used the most" in query:
        return shipping_method_pie_chart(df)
    elif "top selling items" in query:
        return top_selling_products_chart(df)
    elif "top customers by spending" in query:
        return top_customers_by_spending_chart(df)

    else:
        return None

st.title("💬 Business Intelligence Chatbot")
st.write("Hey there! I'm your BI Assistant. Ask me anything about your sales, customers, and trends!")

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

query = st.chat_input("Ask me a business question...")
if query:
    st.session_state.messages.append({"role": "user", "content": query})
    with st.chat_message("user"):
        st.markdown(query)

    query_type = classify_query(query)

    if query_type == "table":
        response = get_table_data(query)
        st.session_state.messages.append({"role": "assistant", "content": str(response)})
        with st.chat_message("assistant"):
            st.markdown(f"**Answer:** {response}")

    elif query_type == "chart":
        chart = get_chart(query)
        if chart:
            st.session_state.messages.append({"role": "assistant", "content": "Here is the requested chart:"})
            with st.chat_message("assistant"):
                st.markdown("Here is the requested chart:")
                st.pyplot(chart)
        else:
            st.session_state.messages.append({"role": "assistant", "content": "Sorry, I couldn't generate the requested chart."})
            with st.chat_message("assistant"):
                st.markdown("Sorry, I couldn't generate the requested chart.")