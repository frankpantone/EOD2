import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime
import sys
import numpy as np

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

print(f"[OK] Pivot table created with {len(pivot_table)} customers and {len(pivot_table.columns)-1} tag types")

# Create PDF
output_file = f"shipment_dashboard_{today.strftime('%Y-%m-%d')}.pdf"

# Set up the PDF with multiple pages
with PdfPages(output_file) as pdf:
    
    # PAGE 1: Title and Key Metrics
    fig = plt.figure(figsize=(11, 8.5))
    fig.patch.set_facecolor('white')
    
    # Title
    plt.text(0.5, 0.95, 'SHIPMENT DASHBOARD', 
             ha='center', va='top', fontsize=32, fontweight='bold',
             color='#2c3e50')
    plt.text(0.5, 0.90, f'Report Date: {today.strftime("%B %d, %Y")}', 
             ha='center', va='top', fontsize=16, color='#7f8c8d')
    
    # Key Metrics Boxes
    metrics = [
        {
            'title': 'SHIPMENTS CREATED TODAY',
            'value': str(total_today),
            'subtitle': f'Date: {today}',
            'color': '#667eea',
            'position': (0.15, 0.75)
        },
        {
            'title': 'TODAY VS TOTAL',
            'value': str(increase),
            'subtitle': f'{increase_pct:.1f}% of total ({total_all} total)',
            'color': '#f5576c',
            'position': (0.55, 0.75)
        },
        {
            'title': 'MOST SHIPPED VEHICLE',
            'value': str(most_shipped_vehicle_count),
            'subtitle': most_shipped_vehicle_name,
            'color': '#00f2fe',
            'position': (0.15, 0.50)
        },
        {
            'title': 'AVG DISTANCE',
            'value': f'{weighted_avg_distance:.0f}',
            'subtitle': 'miles per shipment',
            'color': '#38f9d7',
            'position': (0.55, 0.50)
        }
    ]
    
    for metric in metrics:
        x, y = metric['position']
        # Background box
        rect = mpatches.FancyBboxPatch((x-0.15, y-0.12), 0.3, 0.18,
                                       boxstyle="round,pad=0.01",
                                       facecolor=metric['color'],
                                       edgecolor='none',
                                       alpha=0.9,
                                       transform=fig.transFigure)
        fig.patches.append(rect)
        
        # Text
        plt.text(x, y + 0.04, metric['title'], 
                ha='center', va='center', fontsize=9, fontweight='bold',
                color='white', transform=fig.transFigure)
        plt.text(x, y - 0.02, metric['value'], 
                ha='center', va='center', fontsize=28, fontweight='bold',
                color='white', transform=fig.transFigure)
        plt.text(x, y - 0.08, metric['subtitle'], 
                ha='center', va='center', fontsize=8,
                color='white', transform=fig.transFigure)
    
    # Summary text at bottom
    summary_text = f"""
    Data Summary:
    • Total records processed: {filtered_count} shipments
    • Filtered out {initial_count - filtered_count} records with 'Quote' tag
    • Date range: {df['Created Date'].min().strftime('%m/%d/%Y')} to {df['Created Date'].max().strftime('%m/%d/%Y')}
    • Number of customers: {len(pivot_table)} unique customers
    • Number of tag types: {len(pivot_table.columns)-1}
    """
    
    plt.text(0.5, 0.25, summary_text, 
            ha='center', va='top', fontsize=10,
            color='#2c3e50', transform=fig.transFigure,
            bbox=dict(boxstyle='round', facecolor='#f8f9fa', alpha=0.8, pad=1))
    
    plt.axis('off')
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # PAGE 2: Pivot Table
    fig = plt.figure(figsize=(11, 8.5))
    fig.patch.set_facecolor('white')
    
    plt.text(0.5, 0.96, 'Shipments by Customer and Tag Type', 
             ha='center', va='top', fontsize=18, fontweight='bold',
             color='#2c3e50')
    
    # Prepare table data
    table_data = []
    headers = ['Customer'] + list(pivot_table.columns)
    table_data.append(headers)
    
    for customer, row in pivot_table.iterrows():
        row_data = [customer[:30]]  # Truncate long names
        for col in pivot_table.columns:
            val = row[col]
            row_data.append(str(int(val)) if val > 0 else '')
        table_data.append(row_data)
    
    # Create table
    ax = plt.subplot(111)
    ax.axis('tight')
    ax.axis('off')
    
    table = ax.table(cellText=table_data[1:], 
                    colLabels=table_data[0],
                    cellLoc='center',
                    loc='center',
                    bbox=[0, 0, 1, 0.85])
    
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 2)
    
    # Style header row
    for i in range(len(headers)):
        cell = table[(0, i)]
        cell.set_facecolor('#667eea')
        cell.set_text_props(weight='bold', color='white')
    
    # Alternate row colors and highlight Total column
    for i in range(1, len(table_data)):
        for j in range(len(headers)):
            cell = table[(i, j)]
            if j == len(headers) - 1:  # Total column
                cell.set_facecolor('#e9ecef')
                cell.set_text_props(weight='bold')
            elif i % 2 == 0:
                cell.set_facecolor('#f8f9fa')
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # PAGE 3: Charts
    fig = plt.figure(figsize=(11, 8.5))
    fig.patch.set_facecolor('white')
    
    # Top 10 Customers Bar Chart
    ax1 = plt.subplot(2, 1, 1)
    top_customers = pivot_table.head(10).copy()
    top_customers = top_customers.drop('Total', axis=1)
    
    # Create stacked bar chart
    x_pos = np.arange(len(top_customers))
    bottom = np.zeros(len(top_customers))
    colors = ['#667eea', '#f5576c', '#00f2fe', '#38f9d7', '#ffa502']
    
    for idx, col in enumerate(top_customers.columns):
        values = top_customers[col].values
        ax1.bar(x_pos, values, bottom=bottom, label=col, 
               color=colors[idx % len(colors)], alpha=0.8)
        bottom += values
    
    ax1.set_xlabel('Customer', fontsize=11, fontweight='bold')
    ax1.set_ylabel('Number of Shipments', fontsize=11, fontweight='bold')
    ax1.set_title('Top 10 Customers by Shipment Volume', 
                 fontsize=14, fontweight='bold', pad=20, color='#2c3e50')
    ax1.set_xticks(x_pos)
    ax1.set_xticklabels([name[:20] for name in top_customers.index], 
                        rotation=45, ha='right', fontsize=8)
    ax1.legend(loc='upper right', fontsize=9)
    ax1.grid(axis='y', alpha=0.3)
    
    # Tag Distribution Pie Chart
    ax2 = plt.subplot(2, 1, 2)
    tag_totals = df.groupby('Tags').size().sort_values(ascending=False)
    
    colors_pie = ['#667eea', '#f5576c', '#00f2fe', '#38f9d7', '#ffa502']
    wedges, texts, autotexts = ax2.pie(tag_totals.values, 
                                        labels=tag_totals.index,
                                        autopct='%1.1f%%',
                                        colors=colors_pie[:len(tag_totals)],
                                        startangle=90)
    
    for text in texts:
        text.set_fontsize(10)
        text.set_fontweight('bold')
    
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontsize(9)
        autotext.set_fontweight('bold')
    
    ax2.set_title('Shipment Distribution by Tag Type', 
                 fontsize=14, fontweight='bold', pad=20, color='#2c3e50')
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Set PDF metadata
    d = pdf.infodict()
    d['Title'] = f'Shipment Dashboard - {today}'
    d['Author'] = 'Shipment Reporting System'
    d['Subject'] = 'Daily Shipment Report'
    d['Keywords'] = 'Shipments, Dashboard, Report'
    d['CreationDate'] = datetime.now()

print(f"\n[SUCCESS] PDF Dashboard created successfully: {output_file}")
print(f"\n[METRICS] Key Metrics:")
print(f"   - Shipments Created Today: {total_today}")
print(f"   - Total Shipments (All Time): {total_all}")
print(f"   - Today's Percentage: {increase_pct:.1f}%")
print(f"   - Most Shipped Vehicle: {most_shipped_vehicle_name} ({most_shipped_vehicle_count} units)")
print(f"   - Average Distance: {weighted_avg_distance:.2f} miles")
print(f"\n[INFO] Open the PDF file to view your dashboard report!")






