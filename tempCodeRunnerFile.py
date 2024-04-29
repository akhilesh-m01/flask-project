# Importing necessary libraries
from flask import Flask, render_template, request
import pandas as pd
from datetime import datetime, timedelta

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests
# Initialize Flask application
app = Flask(__name__)

# Function to calculate due dates
def calculate_due_dates(data):
    current_date = datetime.now()
    past_due = data[data['Due Date'] < current_date]
    due_within_7_days = data[(data['Due Date'] >= current_date) & (data['Due Date'] <= current_date + timedelta(days=7))]
    due_from_7_to_15_days = data[(data['Due Date'] > current_date + timedelta(days=7)) & (data['Due Date'] <= current_date + timedelta(days=15))]
    due_from_15_to_30_days = data[(data['Due Date'] > current_date + timedelta(days=15)) & (data['Due Date'] <= current_date + timedelta(days=30))]
    due_more_than_30_days = data[data['Due Date'] > current_date + timedelta(days=30)]

    due_dates = {
        "Due within 7 days": due_within_7_days,
        "Due from 7 to 15 days": due_from_7_to_15_days,
        "Due from 15 to 30 days": due_from_15_to_30_days,
        "Due more than 30 days": due_more_than_30_days
    }

    return due_dates

# Function to generate summary for each period
def generate_summary(period_data):
    summary = []
    vuln_id_counts = period_data.groupby(['Vuln ID', 'Category']).size()
    for (vuln_id, category), count in vuln_id_counts.items():
        summary.append((count, vuln_id, category))
    return summary

# Define the route for uploading Excel file
@app.route('/')
def upload_form():
    return render_template('upload.html')

# Define the route for displaying Excel data and calculating due dates
@app.route('/display', methods=['POST'])
def display_and_calculate():
    # Check if a file is uploaded
    file = request.files['file']
    if file.filename == '':
        return 'No file selected', 400
    
    # Check if the file is in Excel format
    if file and file.filename.endswith(('.xls', '.xlsx')):
        try:
            # Read data from the uploaded Excel file
            data = pd.read_excel(file)
            current_date = datetime.now()

            # Convert 'Due Date' to datetime and filter invalid dates
            data['Due Date'] = pd.to_datetime(data['Due Date'], errors='coerce')
            
             # Get the status selected by the user
            selected_status = request.form.get('status')


           # Filter data based on the selected status
            if selected_status == 'Past Due':
                filtered_data = data[data['PD Flag'] == 'Past Due']
            elif selected_status == 'Not Past Due':
                filtered_data = data[data['PD Flag'] != 'Past Due']
            else:
                return 'Invalid status selected', 400

            if filtered_data.empty:
                return 'No data found for the selected status', 200

            # Check unique values in the "Status" column
            status_unique = filtered_data["Status"].unique()

            # Count occurrences of each unique "Vuln ID" and "Category"
            vuln_id_category_counts = filtered_data.groupby(["Vuln ID", "Category"]).size()

            # Count occurrences of each unique "Status"
            status_counts = filtered_data["Risk Level"].value_counts()

            # Prepare data for HTML table
            output = []
            output.append(status_unique.tolist())
            output.append(f"{selected_status}")
            output.append(f"{len(filtered_data)} Total {status_counts.get('High/Critical', 0)} High, {status_counts.get('Medium', 0)} Medium, {status_counts.get('Low', 0)} Low")

            # Output data for each Vuln ID and Category
            for index, value in vuln_id_category_counts.items():
                vuln_id, category = index
                output.append((value, vuln_id, category, filtered_data.loc[(filtered_data['Vuln ID'] == vuln_id) & (filtered_data['Category'] == category), 'Due Date'].iloc[0]))

            # Calculate due dates for past due items
            due_dates_summary = []
            due_dates = calculate_due_dates(filtered_data)
            for period, period_data in due_dates.items():
                total = len(period_data)
                critical_count = len(period_data[period_data['Risk Level'] == 'Critical'])
                high_count = len(period_data[period_data['Risk Level'] == 'High/Critical'])
                medium_count = len(period_data[period_data['Risk Level'] == 'Medium'])
                low_count = len(period_data[period_data['Risk Level'] == 'Low'])
                due_dates_summary.append((period, total, critical_count, high_count, medium_count, low_count))
                due_dates_summary.extend(generate_summary(period_data))

            return render_template('template.html', data=output, summary_data=due_dates_summary)
        except Exception as e:
            return f'Error reading Excel file: {str(e)}', 500
    else:
        return 'Invalid file format. Please upload an Excel file.', 400

# Run the Flask application
if __name__ == '__main__':
    app.run(debug=True)
