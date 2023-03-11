import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
import time

def within_1_percent(df):
    """
    Filter data within 1% of each other in the Payment Total column and within 5 days of each other.
    """
    if len(df) >= 2:
        mean_payment = df['Payment Total'].mean()
        upper_limit = df['Payment Total'] * 1.01
        lower_limit = df['Payment Total'] * 0.99
        return df[(df['Payment Total'] <= upper_limit) & (df['Payment Total'] >= lower_limit)]
    else:
        return pd.DataFrame()

# Create a root Tkinter window (hidden)
root = tk.Tk()
root.withdraw()

# Ask user to select a file
print("Select a CSV file to load into a Pandas DataFrame")
file_path = filedialog.askopenfilename(filetypes=[('CSV files', '*.csv')])

# Load selected file into Pandas DataFrame
df = pd.read_csv(file_path)

# Convert the Payment Date column to a Pandas DatetimeIndex
df['Payment Date'] = pd.to_datetime(df['Payment Date'])
df['Payment Total'] = pd.to_numeric(df['Payment Total'].str.replace(',', ''), errors='coerce')

# Check for missing or null values in the Payment Total column
print("Number of missing or null values in Payment Total column:", df['Payment Total'].isna().sum())

# Filter data for transactions over 1000
filtered_df = df[df['Payment Total'] > 1000]

# Group by Supplier Name and Payment Date within 1% of each other in the Payment Total column and within 5 days of each other
grouped_df = filtered_df.groupby(['Supplier Name', pd.Grouper(key='Payment Date', freq='5D')]).apply(within_1_percent).reset_index(drop=True)

# Remove empty dataframes from the grouped data
grouped_df = grouped_df[grouped_df['Payment Total'].notna()]

# Create a Tkinter file save dialog box with default extension '.xlsx'
file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])

# Create an Excel writer object
writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

# Write the filtered data to a tab
filtered_df.to_excel(writer, sheet_name='Filtered Data', index=False)

# Write the grouped data to a tab
grouped_df.to_excel(writer, sheet_name='Grouped Data', index=False)

# Create progress bar
total_progress = 100
current_progress = 0
progress_step = total_progress / 5

# Save the Excel file
print("Saving Excel file...")
for i in range(5):
    time.sleep(0.5)  # Artificially slow down saving process  
    writer.save()
    current_progress += progress_step
    print("Progress: {:.0f}%".format(current_progress))

print("Excel file saved successfully!")

# Close the Excel writer object
writer.close()
