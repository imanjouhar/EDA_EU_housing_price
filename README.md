Europe's Shifting Real Estate Market: A Data Story

This is an interactive dashboard built with Dash and Plotly that visualizes the annual rate of change in the House Price Index (HPI) across Europe.

The application is designed to tell a "data story" about the volatility and divergence in European real estate, allowing users to spot regional trends, identify "hot" (positive growth) and "cold" (negative growth) markets, and compare a country's performance against the EU average.

Features

Interactive Choropleth Map: A map of Europe, color-coded with a diverging (Red-Blue) scale to show positive and negative HPI growth for a selected year.

Historical Trend Line: A line chart showing a selected country's HPI change over time, plotted against the EU-27 average for context.

Top 10 Risers: A horizontal bar chart ranking the top 10 countries by HPI growth for the selected year.

Country vs. EU Comparison: A simple bar chart directly comparing a selected country's performance to the EU average for that year.

Full Interactivity (Cross-Filtering): Clicking a country on the map or the "Top 10" chart instantly updates all other graphs to reflect that selection.

Data Source

Provider: Eurostat

Dataset: House Price Index (HPI) - Annual Data [prc_hpi_a]

Local File: The app is hard-coded to read h_pricing_peryear_prc_hpi_a$defaultview_spreadsheet (1).xlsx - Data.csv.

IMPORTANT: This CSV file must be in the same directory as app.py for the application to run.

Setup and Installation

These instructions assume you have Python 3 (e.g., python3.9) installed.

Clone or Download:
Download app.py and the data file h_pricing_peryear_prc_hpi_a$defaultview_spreadsheet (1).xlsx - Data.csv into the same project directory.

Create and Activate a Virtual Environment:
It is highly recommended to use a virtual environment to manage project dependencies.

# Navigate to your project directory
cd path/to/your/project

# Create the environment (e.g., named 'eda_project_workspace')
python3 -m venv eda_project_workspace

# Activate the environment
source eda_project_workspace/bin/activate


Install Required Packages:
With your environment active, install the necessary Python libraries.

# Upgrade pip (recommended)
python -m pip install --upgrade pip

# Install Dash, Pandas, and Plotly
python -m pip install dash pandas plotly


How to Run the Application

Ensure your environment is active:

source eda_project_workspace/bin/activate


Run the Python script:

python app.py


Open in Browser:
The terminal will show an output similar to Running on http://127.0.0.1:8050/. Open this URL in your web browser to view and interact with the dashboard.