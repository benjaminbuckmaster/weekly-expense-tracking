import pandas as pd
from datetime import datetime, timedelta
import pdfkit 
import os

# function to get the date range for the previous week
def previous_week_dates(current_date):
    end_of_last_week = current_date - timedelta(days=current_date.weekday() + 1)
    start_of_last_week = end_of_last_week - timedelta(days=6)
    return start_of_last_week, end_of_last_week

current_date = datetime.now()
start_date, end_date = previous_week_dates(current_date)

xlsx_file = "transactions.xlsx"

df = pd.read_excel(xlsx_file)

# Convert the 'date' column to datetime format
df['date'] = pd.to_datetime(df['date'])

# filter dataframe for the previous weeks transactions
previous_week_df = df[(df['date'] >= start_date) & (df['date'] <= end_date)]

# dictionary to store the category totals
category_totals = {}
other_items = []

# iterrate through the dataframe for previous week
for index, row in previous_week_df.iterrows():
    category = row['category']
    amount = row['amount']
    details = row['details']

    # update the totals for each category
    if category in category_totals:
        category_totals[category] += amount
    else:
        category_totals[category] = amount
    
    # append itemised report for 'other' category to the list
    if category.lower() == 'other':
        other_items.append({'Amount': amount, 'Details': details})

def total_amount(dataframe):
    total_amount = 0

    # iterrate through the dataframe
    for index, row in dataframe.iterrows():
        amount = row['amount']
        total_amount += amount

    return total_amount

# create dataframe for category totals
category_totals_df = pd.DataFrame(list(category_totals.items()), columns=['Category', 'Total'])

# create dataframe for other items
other_items_df = pd.DataFrame(other_items)

# HTML string
html_string = f"""
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        /* CSS styling */

        body {{
            font-family: Arial, sans-serif; /* Use Arial font as fallback */
            display: flex;
            flex-direction: column;
            align-self: center;
            margin: 0 auto; /* Center the body horizontally */
            width: 90%; /* Set the width of the body */
        }}

        table {{
            border-collapse: collapse;
            width: 100%;
            border: 0px solid;
        }}
        th, td {{
            padding: 8px;
            text-align: left;
            border: 1px solid #ddd;
        }}
        th {{
            background-color: rgba(0,0,0,0.1);
        }}

        /* Media queries for smaller screens */
        @media screen and (min-width: 600px) {{
            body {{
                width: 50%; /* Adjust width for smaller screens */
            }}
        }}
    </style>
</head>
<body>
    <h2>Report for week {start_date.strftime("%d-%m-%Y")} to {end_date.strftime("%d-%m-%Y")}</h2>
    <p>Amount spent: ${total_amount(previous_week_df)}</p>
    <h3>Category Totals</h3>
    {category_totals_df.to_html(index=False)}
    <h3>Other Items</h3>
    {other_items_df.to_html(index=False)}
</body>
</html>
"""

# file name for HTML file using the date period
html_file_name = f"reports/report_{start_date.strftime("%d-%m-%Y")}_{end_date.strftime("%d-%m-%Y")}.html"

# write HTML to a file
with open(html_file_name, 'w') as f:
    f.write(html_string)

print(f"\nHTML report has been saved as {html_file_name}\n")