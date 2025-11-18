import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import sys

# Read the CSV file
csv_file = "MB EOD Update_Nov-12-2025-16-16-37.215.csv"

try:
    df = pd.read_csv(csv_file)
    print(f"[OK] Loaded {len(df)} records from {csv_file}")
except Exception as e:
    print(f"Error reading file: {e}")
    sys.exit(1)

# Convert Created Date to datetime
df['Created Date'] = pd.to_datetime(df['Created Date'])

# Rule: Remove orders with tag "CSRM, Quote"
initial_count = len(df)
df = df[~df['Tags'].str.contains('Quote', case=False, na=False)]
filtered_count = len(df)
print(f"[OK] Filtered out {initial_count - filtered_count} records with 'Quote' tag")
print(f"[OK] Working with {filtered_count} records")

# Get today's date (from the data)
today = df['Created Date'].max().date()
print(f"[OK] Latest date in data: {today}")

# Filter for today's shipments
df_today = df[df['Created Date'].dt.date == today]
total_today = len(df_today)

# Calculate total shipments (all dates)
total_all = len(df)

# Calculate increase
increase = total_today
increase_pct = (increase / total_all * 100) if total_all > 0 else 0

# Most shipped vehicle
most_shipped_vehicle = df['Vehicle Info'].value_counts().head(1)
most_shipped_vehicle_name = most_shipped_vehicle.index[0] if len(most_shipped_vehicle) > 0 else "N/A"
most_shipped_vehicle_count = most_shipped_vehicle.values[0] if len(most_shipped_vehicle) > 0 else 0

# Weighted average distance
df['Distance'] = pd.to_numeric(df['Distance'], errors='coerce')
weighted_avg_distance = df['Distance'].mean()

# Create pivot table: Rows = Customer Business Name, Columns = Tags, Values = Count of VIN #
pivot_table = pd.pivot_table(
    df,
    values='VIN #',
    index='Customer Business Name',
    columns='Tags',
    aggfunc='count',
    fill_value=0
)

# Sort by total shipments per customer
pivot_table['Total'] = pivot_table.sum(axis=1)
pivot_table = pivot_table.sort_values('Total', ascending=False)

print(f"\n[OK] Pivot table created with {len(pivot_table)} customers and {len(pivot_table.columns)-1} tag types")

# Create HTML Dashboard
html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Shipment Dashboard - {today}</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1600px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 30px 40px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 600;
        }}
        
        .header .date {{
            font-size: 1.2em;
            opacity: 0.9;
        }}
        
        .metrics {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            padding: 30px 40px;
            background: #f8f9fa;
        }}
        
        .metric-card {{
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        
        .metric-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 15px rgba(0,0,0,0.2);
        }}
        
        .metric-card .label {{
            font-size: 0.9em;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 10px;
            font-weight: 600;
        }}
        
        .metric-card .value {{
            font-size: 2.5em;
            color: #2c3e50;
            font-weight: bold;
            margin-bottom: 5px;
        }}
        
        .metric-card .subvalue {{
            font-size: 1em;
            color: #7f8c8d;
        }}
        
        .metric-card.today {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }}
        
        .metric-card.today .label,
        .metric-card.today .value,
        .metric-card.today .subvalue {{
            color: white;
        }}
        
        .metric-card.increase {{
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            color: white;
        }}
        
        .metric-card.increase .label,
        .metric-card.increase .value,
        .metric-card.increase .subvalue {{
            color: white;
        }}
        
        .metric-card.vehicle {{
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
        }}
        
        .metric-card.vehicle .label,
        .metric-card.vehicle .value,
        .metric-card.vehicle .subvalue {{
            color: white;
        }}
        
        .metric-card.distance {{
            background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);
            color: white;
        }}
        
        .metric-card.distance .label,
        .metric-card.distance .value,
        .metric-card.distance .subvalue {{
            color: white;
        }}
        
        .table-container {{
            padding: 30px 40px;
            overflow-x: auto;
        }}
        
        .table-title {{
            font-size: 1.8em;
            color: #2c3e50;
            margin-bottom: 20px;
            font-weight: 600;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            border-radius: 10px;
            overflow: hidden;
        }}
        
        th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        
        td {{
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
        }}
        
        tr:hover {{
            background-color: #f8f9fa;
        }}
        
        .total-column {{
            font-weight: bold;
            background-color: #e9ecef;
        }}
        
        .chart-container {{
            padding: 30px 40px;
        }}
        
        .chart-title {{
            font-size: 1.8em;
            color: #2c3e50;
            margin-bottom: 20px;
            font-weight: 600;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Shipment Dashboard</h1>
            <div class="date">Report Date: {today.strftime('%B %d, %Y')}</div>
        </div>
        
        <div class="metrics">
            <div class="metric-card today">
                <div class="label">Shipments Created Today</div>
                <div class="value">{total_today}</div>
                <div class="subvalue">Date: {today}</div>
            </div>
            
            <div class="metric-card increase">
                <div class="label">Today vs Total</div>
                <div class="value">{increase}</div>
                <div class="subvalue">{increase_pct:.1f}% of total ({total_all} total)</div>
            </div>
            
            <div class="metric-card vehicle">
                <div class="label">Most Shipped Vehicle</div>
                <div class="value">{most_shipped_vehicle_count}</div>
                <div class="subvalue">{most_shipped_vehicle_name}</div>
            </div>
            
            <div class="metric-card distance">
                <div class="label">Avg Distance</div>
                <div class="value">{weighted_avg_distance:.0f}</div>
                <div class="subvalue">miles per shipment</div>
            </div>
        </div>
        
        <div class="table-container">
            <div class="table-title">Shipments by Customer and Tag Type</div>
            <table>
                <thead>
                    <tr>
                        <th>Customer Business Name</th>
"""

# Add column headers for each tag
for col in pivot_table.columns:
    html_content += f"                        <th>{col}</th>\n"

html_content += """                    </tr>
                </thead>
                <tbody>
"""

# Add table rows
for customer, row in pivot_table.iterrows():
    html_content += f"                    <tr>\n"
    html_content += f"                        <td><strong>{customer}</strong></td>\n"
    for col in pivot_table.columns:
        value = row[col]
        cell_class = 'total-column' if col == 'Total' else ''
        html_content += f"                        <td class='{cell_class}'>{int(value) if value > 0 else ''}</td>\n"
    html_content += f"                    </tr>\n"

html_content += """                </tbody>
            </table>
        </div>
        
        <div class="chart-container">
            <div class="chart-title">Top 10 Customers by Shipment Volume</div>
            <div id="customerChart"></div>
        </div>
        
        <div class="chart-container">
            <div class="chart-title">Shipment Distribution by Tag Type</div>
            <div id="tagChart"></div>
        </div>
    </div>
    
    <script>
"""

# Create chart data for top customers
top_customers = pivot_table.head(10).copy()
top_customers = top_customers.drop('Total', axis=1)

customer_names = list(top_customers.index)
chart_data = []

for col in top_customers.columns:
    chart_data.append({
        'x': customer_names,
        'y': list(top_customers[col].values),
        'name': col,
        'type': 'bar'
    })

html_content += f"""
        var customerData = {chart_data};
        
        var customerLayout = {{
            barmode: 'stack',
            height: 500,
            xaxis: {{
                tickangle: -45
            }},
            yaxis: {{
                title: 'Number of Shipments'
            }},
            margin: {{
                b: 150
            }}
        }};
        
        Plotly.newPlot('customerChart', customerData, customerLayout);
"""

# Create pie chart for tag distribution
tag_totals = df.groupby('Tags').size().sort_values(ascending=False)
tag_names = list(tag_totals.index)
tag_values = list(tag_totals.values)

html_content += f"""
        var tagData = [{{
            values: {tag_values},
            labels: {tag_names},
            type: 'pie',
            textinfo: 'label+percent',
            textposition: 'auto',
            hovertemplate: '<b>%{{label}}</b><br>Count: %{{value}}<br>Percentage: %{{percent}}<extra></extra>'
        }}];
        
        var tagLayout = {{
            height: 500,
            showlegend: true
        }};
        
        Plotly.newPlot('tagChart', tagData, tagLayout);
    </script>
</body>
</html>
"""

# Save the HTML file
output_file = f"shipment_dashboard_{today.strftime('%Y-%m-%d')}.html"
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\n[SUCCESS] Dashboard created successfully: {output_file}")
print(f"\n[METRICS] Key Metrics:")
print(f"   - Shipments Created Today: {total_today}")
print(f"   - Total Shipments (All Time): {total_all}")
print(f"   - Today's Percentage: {increase_pct:.1f}%")
print(f"   - Most Shipped Vehicle: {most_shipped_vehicle_name} ({most_shipped_vehicle_count} units)")
print(f"   - Average Distance: {weighted_avg_distance:.2f} miles")
print(f"\n[INFO] Open the HTML file in your browser to view the interactive dashboard!")

