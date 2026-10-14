from flask import Flask, request, send_file
import pandas as pd
import plotly.express as px
from io import BytesIO
from html import escape

app = Flask(__name__)

EXCEL_FILE = "blinkit_analysis.xlsx"


@app.route("/download")
def download_excel():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Analysis")

    # Apply the same search and filters
    search = request.args.get("search", "").strip()
    category = request.args.get("category", "")
    delivery = request.args.get("delivery", "")
    rating = request.args.get("rating", "")

    if search:
        df = df[
            df["order_id"].astype(str).str.contains(
                search, case=False, na=False, regex=False
            )
            | df["product_name"].astype(str).str.contains(
                search, case=False, na=False, regex=False
            )
        ]

    if category:
        df = df[df["category"].astype(str) == category]

    if delivery:
        df = df[df["delivery_status"].astype(str) == delivery]

    if rating in ["1", "2", "3", "4", "5"]:
        numeric_rating = pd.to_numeric(
            df["rating"], errors="coerce"
        )
        df = df[numeric_rating == int(rating)]

    # Create Excel file in memory
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Analysis")

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Blinkit_Filtered_Data.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/")
def home():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Analysis")

    # Get search and filter values
    search = request.args.get("search", "").strip()
    category = request.args.get("category", "")
    delivery = request.args.get("delivery", "")
    rating = request.args.get("rating", "")

    # Save original data for dropdown options
    all_categories = sorted(
        df["category"].dropna().astype(str).unique()
    )
    all_deliveries = sorted(
        df["delivery_status"].dropna().astype(str).unique()
    )

    # Search by Order ID or Product Name
    if search:
        df = df[
            df["order_id"].astype(str).str.contains(
                search, case=False, na=False, regex=False
            )
            | df["product_name"].astype(str).str.contains(
                search, case=False, na=False, regex=False
            )
        ]

    # Category filter
    if category in all_categories:
        df = df[df["category"].astype(str) == category]

    # Delivery status filter
    if delivery in all_deliveries:
        df = df[df["delivery_status"].astype(str) == delivery]

    # Rating filter
    if rating in ["1", "2", "3", "4", "5"]:
        numeric_rating = pd.to_numeric(
            df["rating"], errors="coerce"
        )
        df = df[numeric_rating == int(rating)]

    # Dashboard values based on filtered data
    total_sales = pd.to_numeric(
        df["order_total"], errors="coerce"
    ).sum()

    total_orders = df["order_id"].nunique()

    avg_order_value = (
        total_sales / total_orders if total_orders else 0
    )

    avg_rating = pd.to_numeric(
        df["rating"], errors="coerce"
    ).mean()

    rating_text = (
        f"{avg_rating:.2f} / 5"
        if pd.notna(avg_rating) else "N/A"
    )



    # Prepare data for charts
    chart_df = df.copy()

    chart_df["order_total"] = pd.to_numeric(
        chart_df["order_total"], errors="coerce"
    )

    chart_df["rating"] = pd.to_numeric(
        chart_df["rating"], errors="coerce"
    )

    chart_df["order_date"] = pd.to_datetime(
        chart_df["order_date"], errors="coerce"
    )

    # 1. Monthly Sales Trend
    monthly = chart_df.dropna(
        subset=["order_date", "order_total"]
    ).copy()

    monthly["Month"] = (
        monthly["order_date"].dt.to_period("M").astype(str)
    )

    monthly_sales = monthly.groupby(
        "Month", as_index=False
    )["order_total"].sum()

    fig1 = px.line(
        monthly_sales,
        x="Month",
        y="order_total",
        markers=True,
        title="Monthly Sales Trend",
        labels={"order_total": "Sales (₹)", "Month": "Month"}
    )

    # 2. Category-wise Sales
    category_sales = chart_df.groupby(
        "category", as_index=False
    )["order_total"].sum()

    fig2 = px.bar(
        category_sales,
        x="category",
        y="order_total",
        title="Category-wise Sales",
        labels={"category": "Category", "order_total": "Sales (₹)"}
    )

    # 3. Payment Method-wise Sales
    payment_sales = chart_df.groupby(
        "payment_method", as_index=False
    )["order_total"].sum()

    fig3 = px.bar(
        payment_sales,
        x="payment_method",
        y="order_total",
        title="Payment Method-wise Sales",
        labels={
            "payment_method": "Payment Method",
            "order_total": "Sales (₹)"
        }
    )

    # 4. Rating-wise Customer Feedback
    rating_data = chart_df.dropna(subset=["rating"]).copy()
    rating_data["Rating"] = (
        rating_data["rating"].astype(int).astype(str) + " Star"
    )

    rating_counts = rating_data.groupby(
        "Rating", as_index=False
    ).size().rename(columns={"size": "Customers"})

    fig4 = px.bar(
        rating_counts,
        x="Rating",
        y="Customers",
        title="Rating-wise Customer Feedback"
    )

    # 5. Sentiment-wise Customer Feedback
    sentiment_data = chart_df.dropna(
        subset=["sentiment"]
    )

    sentiment_counts = sentiment_data.groupby(
        "sentiment", as_index=False
    ).size().rename(columns={"size": "Records"})

    fig5 = px.pie(
        sentiment_counts,
        names="sentiment",
        values="Records",
        title="Sentiment-wise Customer Feedback"
    )

    # Convert charts into website HTML
    chart1 = fig1.to_html(
        full_html=False, include_plotlyjs="cdn"
    )
    chart2 = fig2.to_html(
        full_html=False, include_plotlyjs=False
    )
    chart3 = fig3.to_html(
        full_html=False, include_plotlyjs=False
    )
    chart4 = fig4.to_html(
        full_html=False, include_plotlyjs=False
    )
    chart5 = fig5.to_html(
        full_html=False, include_plotlyjs=False
    )

    # Prepare table data
    columns = [
        "order_id", "order_date", "product_name", "category",
        "brand", "quantity", "unit_price", "order_total",
        "payment_method", "delivery_status", "rating", "sentiment"
    ]

    table_df = df[columns].copy().fillna("")

    # Safely display Excel values in HTML
    table_html = table_df.to_html(
        index=False,
        border=0,
        classes="data-table",
        escape=True
    )

    # Create dropdown options
    category_options = '<option value="">All Categories</option>'
    for item in all_categories:
        selected = "selected" if item == category else ""
        category_options += (
            f'<option value="{escape(item, quote=True)}" {selected}>'
            f'{escape(item)}</option>'
        )

    delivery_options = '<option value="">All Delivery Statuses</option>'
    for item in all_deliveries:
        selected = "selected" if item == delivery else ""
        delivery_options += (
            f'<option value="{escape(item, quote=True)}" {selected}>'
            f'{escape(item)}</option>'
        )

    rating_options = '<option value="">All Ratings</option>'
    for item in ["1", "2", "3", "4", "5"]:
        selected = "selected" if item == rating else ""
        rating_options += (
            f'<option value="{item}" {selected}>{item} Star</option>'
        )

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Blinkit Sales Dashboard</title>

        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 30px;
                background: #f4f7f4;
                color: #333;
            }}

            h1 {{ color: #16803c; }}

            .cards {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 20px;
                margin-top: 25px;
            }}

            .card {{
                background: white;
                padding: 25px;
                border-radius: 12px;
                box-shadow: 0 3px 10px #00000012;
                border-left: 5px solid #16803c;
            }}

            .card h3 {{
                color: #555;
                margin-top: 0;
                font-size: 16px;
            }}

            .card p {{
                color: #16803c;
                font-size: 27px;
                font-weight: bold;
                margin-bottom: 0;
            }}

            .filters {{
                background: white;
                padding: 20px;
                margin-top: 30px;
                border-radius: 12px;
                box-shadow: 0 3px 10px #00000012;
            }}

            .filter-row {{
                display: flex;
                flex-wrap: wrap;
                gap: 12px;
                margin-top: 15px;
            }}

            input, select, button, .reset {{
                padding: 11px;
                border: 1px solid #ccc;
                border-radius: 6px;
                font-size: 14px;
            }}

            input {{ flex: 1; min-width: 180px; }}

            button {{
                background: #16803c;
                color: white;
                border: none;
                cursor: pointer;
            }}

            .reset {{
                background: #eee;
                color: #333;
                text-decoration: none;
            }}

            .table-container {{
                overflow-x: auto;
                background: white;
                margin-top: 20px;
                border-radius: 10px;
            }}

            .data-table {{
                width: 100%;
                border-collapse: collapse;
                white-space: nowrap;
            }}

            .data-table th, .data-table td {{
                padding: 12px;
                text-align: left;
                border-bottom: 1px solid #ddd;
            }}

            .data-table th {{
                background: #16803c;
                color: white;
            }}

            .data-table tr:hover {{ background: #f1f8f1; }}

            .note {{ color: #666; margin-top: 25px; }}




            .charts-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
    gap: 20px;
    margin-top: 20px;
}}

        .chart-card {{
         background: white;
         padding: 12px;
         border-radius: 12px;
         box-shadow: 0 3px 10px #00000012;
         min-width: 0;
        overflow: hidden;
         }}
        </style>
    </head>

    <body>
        <h1>Blinkit Sales Dashboard</h1>
        <p>Sales and Customer Analysis</p>

        <div class="cards">
            <div class="card">
                <h3>Total Sales</h3>
                <p>₹{total_sales:,.2f}</p>
            </div>

            <div class="card">
                <h3>Total Orders</h3>
                <p>{total_orders:,}</p>
            </div>

            <div class="card">
                <h3>Average Order Value</h3>
                <p>₹{avg_order_value:,.2f}</p>
            </div>

            <div class="card">
                <h3>Average Rating</h3>
                <p>{rating_text}</p>
            </div>
        </div>

        <div class="filters">
            <h2>Search and Filters</h2>

            <form method="GET" action="/">
                <div class="filter-row">
                    <input
                        type="text"
                        name="search"
                        placeholder="Search Order ID or Product Name"
                        value="{escape(search, quote=True)}"
                    >

                    <select name="category">
                        {category_options}
                    </select>

                    <select name="delivery">
                        {delivery_options}
                    </select>

                    <select name="rating">
                        {rating_options}
                    </select>

                    <button type="submit">Search / Apply Filters</button>
                    <a class="reset"
   href="/download?search={escape(search, quote=True)}&category={escape(category, quote=True)}&delivery={escape(delivery, quote=True)}&rating={escape(rating, quote=True)}">
   Download Excel
</a>
                    <a class="reset" href="/">Reset</a>
                </div>
            </form>
        </div>
        <h2>Sales & Customer Analytics</h2>

    <div class="charts-grid">
        <div class="chart-card">{chart1}</div>
        <div class="chart-card">{chart2}</div>
        <div class="chart-card">{chart3}</div>
        <div class="chart-card">{chart4}</div>
        <div class="chart-card">{chart5}</div>
    </div>
        <h2>Order Details</h2>
        <p>Matching records: {len(df):,}</p>

        <div class="table-container">
            {table_html}
        </div>

        <p class="note">
            Dashboard values and order records are based on your Excel data.
        </p>
    </body>
    </html>
    """


if __name__ == "__main__":
    app.run(debug=True)
