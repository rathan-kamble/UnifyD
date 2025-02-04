import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Simulated Data Sources

# 1. Simulated Outlook Data (Email)
outlook_data = pd.DataFrame({
    'Email': ['customer1@example.com', 'customer2@example.com'],
    'Engagement': ['Follow-up', 'Proposal Sent'],
    'Date': ['2023-12-01', '2023-12-05']
})

# 2. Simulated SharePoint Data
sharepoint_data = pd.DataFrame({
    'Region': ['North', 'South'],
    'Salesperson': ['Alice', 'Bob'],
    'Amount': [50000, 40000]
})

# 3. Simulated Salesforce Data
salesforce_data = pd.DataFrame({
    'Opportunity Name': ['Deal A', 'Deal B'],
    'Account': ['Company X', 'Company Y'],
    'Amount': [70000, 30000],
    'Stage': ['Closed Won', 'Closed Won']
})

# 4. Simulated Excel Data
excel_data = pd.DataFrame({
    'Region': ['East', 'West'],
    'Salesperson': ['Charlie', 'Dana'],
    'Amount': [60000, 45000]
})

# Combine Data Sources

# Add consistent columns where necessary
outlook_data['Source'] = 'Outlook'
sharepoint_data['Source'] = 'SharePoint'
salesforce_data['Source'] = 'Salesforce'
excel_data['Source'] = 'Excel'

# Consolidate all data into a single DataFrame
final_data = pd.concat([
    outlook_data,
    sharepoint_data.rename(columns={'Amount': 'Sales Amount'}),
    salesforce_data.rename(columns={'Amount': 'Sales Amount'}),
    excel_data.rename(columns={'Amount': 'Sales Amount'})
], ignore_index=True)

# Standardize Date Format
final_data['Date'] = final_data.get('Date', pd.NaT)
final_data['Date'] = pd.to_datetime(final_data['Date'], errors='coerce')

# Save Raw Data to Excel
final_data.to_excel('Consolidated_Raw_Data.xlsx', index=False)
print("Consolidated raw data saved to 'Consolidated_Raw_Data.xlsx'.")

# Data Analysis

# Group Sales Amount by Region
sales_summary = final_data.groupby('Region', as_index=False).agg({'Sales Amount': 'sum'})
print("Sales Summary by Region:\n", sales_summary)

# Save Sales Summary to Excel
sales_summary.to_excel('Sales_Summary.xlsx', index=False)
print("Sales summary saved to 'Sales_Summary.xlsx'.")

# Visualization
plt.figure(figsize=(8, 5))
plt.bar(sales_summary['Region'], sales_summary['Sales Amount'], color=['blue', 'green', 'orange', 'red'])
plt.title('Sales Performance by Region', fontsize=14)
plt.xlabel('Region', fontsize=12)
plt.ylabel('Sales Amount', fontsize=12)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('Sales_Performance_Chart.png')
plt.show()
print("Sales performance chart saved to 'Sales_Performance_Chart.png'.")

# End of Script
