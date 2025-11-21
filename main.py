import dash
from dash import dcc, html, Input, Output, State, callback_context
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import numpy as np
import os
import warnings

# Suppress warnings to keep the console output clean
warnings.simplefilter("ignore")

# ==========================================
# CONFIGURATION
# ==========================================

# This variable holds the absolute path to the Excel file containing the data.
DATA_FILE = r'/Users/imanjouhar/Documents/UNI DOCS/EXPLORATIVE DATA ANALYSIS/code EDA PROJECT UNI/house_pricing.xlsx'

# ==========================================
# DATA PROCESSING
# ==========================================

# ---------------------------------------------------------
# FUNCTION: load_and_clean_data
# ---------------------------------------------------------
# This is the main ETL (Extract, Transform, Load) function.
# It performs the following steps:
# 1. Checks if the file exists.
# 2. Loads the raw Excel sheet named 'Data'.
# 3. Identifies the starting rows for three different datasets within the single sheet.
# 4. Cleans, formats, and melts (pivots) the data into a usable structure.
# 5. Maps country names to ISO-3 codes for mapping purposes.
# ---------------------------------------------------------
def load_and_clean_data(file_path):
    # Check if file exists on the system
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        # Load the file using openpyxl engine
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        if 'Data' in xls.sheet_names:
            df_raw = pd.read_excel(xls, sheet_name='Data', header=None)
        else:
            raise ValueError("Sheet 'Data' not found in Excel file.")
            
    except Exception as e:
        raise ValueError(f"Could not open Excel file: {e}")

    # --- HELPER: FIND ROW INDICES ---
    # This internal function searches the raw dataframe for a specific string (title)
    # to determine where a data table begins. It looks ahead 15 rows to find the "TIME" header.
    def find_start_row(df, search_term):
        matches = df[df.apply(lambda row: row.astype(str).str.contains(search_term).any(), axis=1)]
        if not matches.empty:
            idx = matches.index[0]
            for i in range(idx, idx + 15):
                if "TIME" in str(df.iloc[i].values):
                    return i
        return -1

    # locate the start rows for the three categories of data
    row_new = find_start_row(df_raw, "Purchases of newly built dwellings")
    row_exist = find_start_row(df_raw, "Purchases of existing dwellings")
    row_total = find_start_row(df_raw, "Total")
    
    # Fallback logic: If 'Total' isn't found by name, look for the first occurrence of 'TIME'
    if row_total == -1: 
        time_rows = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains('TIME').any(), axis=1)].index
        if len(time_rows) > 0: row_total = time_rows[0]

    # --- HELPER: EXTRACT AND CLEAN BLOCK ---
    # This internal function takes a starting row index, slices the dataframe,
    # sets the correct headers, and converts the table from "Wide" format (years as columns)
    # to "Long" format (Year as a column value) for easier plotting.
    def extract_block(start_row, label_type):
        if start_row == -1: return pd.DataFrame()
        
        # Slice the dataframe starting from the found row
        chunk = df_raw.iloc[start_row:]
        chunk.columns = chunk.iloc[0] # Set first row as header
        chunk = chunk[1:] # Remove the header row from data
        chunk.rename(columns={chunk.columns[0]: 'Country'}, inplace=True)
        
        # Identify columns that look like years (4 digits)
        year_cols = [c for c in chunk.columns if str(c).strip().replace('.0','').isdigit() and len(str(c).strip()) == 4]
        
        if not year_cols: return pd.DataFrame()

        # Melt the dataframe (pivot)
        df_long = pd.melt(chunk, id_vars=['Country'], value_vars=year_cols, var_name='Year', value_name='Value')
        
        # Ensure numeric types
        df_long['Value'] = pd.to_numeric(df_long['Value'], errors='coerce')
        df_long['Year'] = pd.to_numeric(df_long['Year'].astype(str).str.replace('.0', '', regex=False))
        df_long.dropna(subset=['Country'], inplace=True)
        df_long['Country'] = df_long['Country'].astype(str).str.strip()
        df_long['Type'] = label_type
        
        return df_long

    # Extract the three dataframes using the helper function
    df_total = extract_block(row_total, 'Total')
    df_new = extract_block(row_new, 'New')
    df_exist = extract_block(row_exist, 'Existing')

    # Filter Countries (List of countries to keep in the analysis)
    countries_to_keep = [
        'Belgium', 'Bulgaria', 'Czechia', 'Denmark', 'Germany', 'Estonia', 'Ireland', 
        'Greece', 'Spain', 'France', 'Croatia', 'Italy', 'Cyprus', 'Latvia', 'Lithuania', 
        'Luxembourg', 'Hungary', 'Malta', 'Netherlands', 'Austria', 'Poland', 'Portugal', 
        'Romania', 'Slovenia', 'Slovakia', 'Finland', 'Sweden', 'Iceland', 'Norway'
    ] 
    
    # Internal function to filter the dataframe to only include the desired countries and EU aggregates
    def clean_final(df):
        if df.empty: return df
        mask = df['Country'].isin(countries_to_keep)
        aggregates_search = ['European Union', 'EU27', 'Euro area']
        mask_agg = df['Country'].apply(lambda x: any(agg in str(x) for agg in aggregates_search))
        return df[mask | mask_agg].copy()

    # Apply filtering
    df_total = clean_final(df_total)
    df_new = clean_final(df_new)
    df_exist = clean_final(df_exist)

    # Map country names to ISO 3-letter codes (Required for Plotly Maps)
    iso_map = {
        'Greece': 'GRC', 'EL': 'GRC', 'United Kingdom': 'GBR', 'UK': 'GBR',
        'Belgium': 'BEL', 'Bulgaria': 'BGR', 'Czechia': 'CZE', 'Denmark': 'DNK',
        'Germany': 'DEU', 'Estonia': 'EST', 'Ireland': 'IRL', 'Spain': 'ESP',
        'France': 'FRA', 'Croatia': 'HRV', 'Italy': 'ITA', 'Cyprus': 'CYP',
        'Latvia': 'LVA', 'Lithuania': 'LTU', 'Luxembourg': 'LUX', 'Hungary': 'HUN',
        'Malta': 'MLT', 'Netherlands': 'NLD', 'Austria': 'AUT', 'Poland': 'POL',
        'Portugal': 'PRT', 'Romania': 'ROU', 'Slovenia': 'SVN', 'Slovakia': 'SVK',
        'Finland': 'FIN', 'Sweden': 'SWE', 'Iceland': 'ISL', 'Norway': 'NOR',
        'Switzerland': 'CHE'
    }
    
    df_total['iso_alpha'] = df_total['Country'].map(iso_map)
    # Manual fixes for cases that might have been missed
    df_total.loc[df_total['iso_alpha'].isna() & df_total['Country'].str.contains('Greece'), 'iso_alpha'] = 'GRC'
    df_total.loc[df_total['iso_alpha'].isna() & df_total['Country'].str.contains('Kingdom'), 'iso_alpha'] = 'GBR'

    return df_total, df_new, df_exist

# ==========================================
# INITIALIZATION
# ==========================================
# This block attempts to load the data immediately when the script starts.
# It separates the data into "Individual Countries" and "EU Aggregates".
# It creates empty placeholders if the data loading fails to prevent the app from crashing.
try:
    df_total, df_new, df_exist = load_and_clean_data(DATA_FILE)
    AVAILABLE_YEARS = sorted(df_total['Year'].unique())
    MIN_YEAR, MAX_YEAR = min(AVAILABLE_YEARS), max(AVAILABLE_YEARS)
    
    aggregates_search = ['European Union', 'EU27', 'Euro area']
    mask_eu = df_total['Country'].apply(lambda x: any(agg in str(x) for agg in aggregates_search))
    df_eu_total = df_total[mask_eu].copy()
    df_countries_total = df_total[~mask_eu & df_total['iso_alpha'].notna()].copy()
    
except Exception as e:
    print(f"Data Load Error: {e}")
    # Create empty objects so the app can still launch (even if blank)
    df_countries_total = pd.DataFrame()
    df_eu_total = pd.DataFrame()
    df_new, df_exist = pd.DataFrame(), pd.DataFrame()
    MIN_YEAR, MAX_YEAR = 2015, 2024
    AVAILABLE_YEARS = []

# ==========================================
# APP LAYOUT
# ==========================================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.LITERA])

# Styling for the cards
CARD_STYLE = {
    "boxShadow": "0 2px 4px 0 rgba(0,0,0,0.1)", 
    "borderRadius": "5px", 
    "marginBottom": "0px"
}

app.layout = dbc.Container([
    # Store: Keeps the selected country in the browser's memory (Client-side storage)
    dcc.Store(id='store-country', data='Germany'),

    # --- Header & Slider ---
    dbc.Row([
        dbc.Col([
            html.H2("European Real Estate Dashboard", className="fw-bold mb-0"),
            html.Small("Annual House Price Rate of Change (%)", className="text-muted")
        ], width=4, className="d-flex flex-column justify-content-center"),

        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.Label("Select Year:", className="fw-bold mb-0"),
                    dcc.Slider(
                        id='year-slider', min=MIN_YEAR, max=MAX_YEAR, value=MAX_YEAR,
                        marks={str(y): str(y) for y in AVAILABLE_YEARS}, step=None,
                        className="p-0"
                    )
                ], className="p-2") 
            ], style=CARD_STYLE)
        ], width=8)
    ], className="my-2 align-items-center"),

    # --- Main Content Area ---
    dbc.Row([
        # LEFT COLUMN: THE BIG MAP (8/12 width)
        dbc.Col([
            dbc.Card([
                dbc.CardHeader("Geographic Heatmap (Total HPI)", className="py-2 fw-bold"),
                dbc.CardBody(
                    dcc.Graph(id='map-graph', style={'height': '700px'}), 
                    className="p-0"
                )
            ], style=CARD_STYLE, className="h-100")
        ], lg=8, className="pe-1"),

        # RIGHT COLUMN: TABS (4/12 width)
        dbc.Col([
            dbc.Card([
                dbc.CardHeader(
                    dcc.Tabs(id='analysis-tabs', value='tab-ranking', children=[
                        dcc.Tab(label='Rankings', value='tab-ranking', className="fw-bold"),
                        dcc.Tab(label='Market Overview', value='tab-overview', className="fw-bold"),
                        dcc.Tab(label='Sector Analysis', value='tab-sectors', className="fw-bold"),
                    ], colors={"border": "#d6d6d6", "primary": "#e67e22", "background": "#f8f9fa"}),
                    className="p-0 border-bottom-0"
                ),
                dbc.CardBody(
                    html.Div(id='tabs-content'), 
                    className="p-1",
                    style={'height': '700px'} 
                )
            ], style=CARD_STYLE, className="h-100")
        ], lg=4, className="ps-1")

    ], className="g-0"),

], fluid=True, style={'backgroundColor': '#f8f9fa', 'minHeight': '100vh', 'padding': '10px'})

# ==========================================
# CALLBACKS
# ==========================================

# ---------------------------------------------------------
# FUNCTION: update_selection
# ---------------------------------------------------------
# This callback handles user interaction with the map.
# Input: A click event on the 'map-graph'.
# Output: Updates the 'store-country' component with the name of the clicked country.
# If no click happens (or on initial load), it preserves the current state.
# ---------------------------------------------------------
@app.callback(
    Output('store-country', 'data'),
    Input('map-graph', 'clickData'),
    State('store-country', 'data')
)
def update_selection(map_clk, current):
    ctx = callback_context
    if not ctx.triggered: return current
    if map_clk:
        # Extract the ISO code from the clicked point
        iso = map_clk['points'][0]['location']
        # Match ISO code to Country Name
        found = df_countries_total[df_countries_total['iso_alpha'] == iso]
        if not found.empty: return found['Country'].iloc[0]
    return current

# ---------------------------------------------------------
# FUNCTION: update_map
# ---------------------------------------------------------
# This callback renders the main Choropleth map.
# Inputs: The selected Year (slider) and the selected Country (store).
# Output: A Plotly Figure object containing the map.
# It handles color scaling logic to ensure 0% is white, negative is blue, positive is red.
# It also adds a highlighting border around the selected country.
# ---------------------------------------------------------
@app.callback(
    Output('map-graph', 'figure'),
    Input('year-slider', 'value'),
    Input('store-country', 'data')
)
def update_map(year, country):
    if df_countries_total.empty: return go.Figure()

    dff = df_countries_total[df_countries_total['Year'] == year].copy()
    
    # --- CALIBRATED COLOR SCALE ---
    # Range: -10 to 25 (Span = 35)
    min_val, max_val = -10.0, 25.0
    span = max_val - min_val
    
    # Calculate normalized positions (0.0 to 1.0) for the color gradients
    # We want vivid colors starting at +/- 2% from zero.
    norm_neg_2 = ( -2.0 - min_val ) / span # ~0.228
    norm_zero  = (  0.0 - min_val ) / span # ~0.286
    norm_pos_2 = (  2.0 - min_val ) / span # ~0.343

    # Custom Diverging scale
    custom_scale = [
        [0.0, '#313695'],         # -10: Deep Blue
        [norm_neg_2, "#b3f0f4"],  #  -2: Distinct Light Blue
        [norm_zero, '#ffffff'],   #   0: Pure White
        [norm_pos_2, "#f4f4a7"],  #  +2: Distinct Light Yellow 
        [1.0, '#a50026']          # +25: Deep Red
    ]

    # Create the base map
    fig_map = px.choropleth(
        dff, locations='iso_alpha', color='Value',
        hover_name='Country',
        color_continuous_scale=custom_scale, 
        range_color=[min_val, max_val], 
    )
    
    # Update layout for visual aesthetics (removing margins, setting projection)
    fig_map.update_layout(
        margin={"r":0,"t":0,"l":0,"b":0}, 
        geo_bgcolor='rgba(0,0,0,0)',
        geo=dict(
            scope='europe',        
            projection_scale=0.9,  
            center=dict(lat=54, lon=15), 
            resolution=50
        )
    )

    # Add a specific trace to highlight the selected country (Border effect)
    dff_highlight = df_countries_total[df_countries_total['Country'] == country]
    if not dff_highlight.empty:
        iso = dff_highlight['iso_alpha'].iloc[0]
        fig_map.add_trace(go.Choropleth(
            locations=[iso], z=[1], locationmode='ISO-3',
            colorscale=[[0,'rgba(0,0,0,0)'],[1,'rgba(0,0,0,0)']], showscale=False,
            marker_line_color='black', marker_line_width=3
        ))
    return fig_map

# ---------------------------------------------------------
# FUNCTION: render_content
# ---------------------------------------------------------
# This callback controls the right-hand side panel content.
# Inputs: The selected Tab ('Rankings', 'Overview', or 'Sectors'), Year, and Country.
# Output: The specific graphs or HTML content corresponding to the selected tab.
# ---------------------------------------------------------
@app.callback(
    Output('tabs-content', 'children'),
    [Input('analysis-tabs', 'value'),
     Input('year-slider', 'value'),
     Input('store-country', 'data')]
)
def render_content(tab, year, country):
    if df_countries_total.empty: return html.Div("No Data")

    # CASE 1: RANKING TAB (Lollipop Chart)
    if tab == 'tab-ranking':
        dff_lollipop = df_countries_total[df_countries_total['Year'] == year].copy()
        dff_lollipop = dff_lollipop.drop_duplicates(subset=['Country'], keep='last')
        df_sorted = dff_lollipop.sort_values('Value', ascending=True)

        # Highlight selected country in Orange, others in Green
        cols_markers = ['#2ecc71'] * len(df_sorted)
        if country in df_sorted['Country'].values:
            cols_markers[list(df_sorted['Country']).index(country)] = '#e67e22' 

        fig_risers = go.Figure()
        
        # Draw lines (the stick of the lollipop)
        for i, row in df_sorted.iterrows():
            fig_risers.add_trace(go.Scatter(
                x=[0, row['Value']], y=[row['Country'], row['Country']],
                mode='lines', line=dict(color='#3498db', width=2),
                showlegend=False, hoverinfo='skip'
            ))
        
        # Draw markers (the candy of the lollipop)
        fig_risers.add_trace(go.Scatter(
            x=df_sorted['Value'], y=df_sorted['Country'],
            mode='markers', marker=dict(size=8, color=cols_markers),
            showlegend=False, hoverinfo='x+y'
        ))

        fig_risers.update_layout(
            title=f"Market Ranking ({year})",
            template='plotly_white', 
            xaxis=dict(title='Change (%)', zeroline=True, showgrid=True),
            yaxis=dict(type='category', dtick=1, showgrid=True),
            margin=dict(l=0, r=10, t=40, b=30), 
            height=680 
        )
        return dcc.Graph(figure=fig_risers, style={'height': '100%'})

    # CASE 2: MARKET OVERVIEW TAB (Line Chart + Range Area)
    elif tab == 'tab-overview':
        # Calculate statistics across all countries per year
        avg_series = df_countries_total.groupby('Year')['Value'].mean()
        min_series = df_countries_total.groupby('Year')['Value'].min()
        max_series = df_countries_total.groupby('Year')['Value'].max()
        
        fig_overview = go.Figure()
        
        # Plot Max line (hidden line, used for fill)
        fig_overview.add_trace(go.Scatter(
            x=avg_series.index, y=max_series,
            mode='lines', line=dict(width=0),
            showlegend=False, hoverinfo='skip'
        ))
        # Plot Min line and fill area up to Max line
        fig_overview.add_trace(go.Scatter(
            x=avg_series.index, y=min_series,
            mode='lines', line=dict(width=0),
            fill='tonexty', fillcolor='rgba(0,100,80,0.2)',
            name='Market Range (Min-Max)'
        ))
        # Plot Average line
        fig_overview.add_trace(go.Scatter(
            x=avg_series.index, y=avg_series,
            mode='lines+markers',
            name='Average of All Countries',
            line=dict(color='#e67e22', width=3)
        ))
        # Add vertical line indicating currently selected year
        fig_overview.add_vline(x=year, line_dash="dash", line_color="black")

        fig_overview.update_layout(
            title="General Market Growth (All Countries Average)",
            template='plotly_white',
            yaxis=dict(title='Average Change (%)'),
            xaxis=dict(title='Year', dtick=1),
            margin=dict(l=0, r=10, t=40, b=30),
            height=680,
            legend=dict(orientation="h", y=1.02)
        )
        return dcc.Graph(figure=fig_overview, style={'height': '100%'})

    # CASE 3: SECTOR ANALYSIS TAB (Line Chart + Box Plot)
    elif tab == 'tab-sectors':
        # Filter data for the selected country only
        ts_total = df_countries_total[df_countries_total['Country'] == country].sort_values('Year')
        ts_new = df_new[df_new['Country'] == country].sort_values('Year')
        ts_exist = df_exist[df_exist['Country'] == country].sort_values('Year')

        # Create Line Chart (Trends over time)
        fig_line = go.Figure()
        fig_line.add_trace(go.Scatter(x=ts_total['Year'], y=ts_total['Value'], name='Total', line=dict(color='black', width=3)))
        fig_line.add_trace(go.Scatter(x=ts_new['Year'], y=ts_new['Value'], name='New Dwellings', line=dict(color='#2ecc71', width=2)))
        fig_line.add_trace(go.Scatter(x=ts_exist['Year'], y=ts_exist['Value'], name='Existing Dwellings', line=dict(color='#e74c3c', width=2, dash='dot')))

        fig_line.update_layout(
            title=f"Sector Trends: {country}",
            template='plotly_white',
            hovermode="x unified",
            legend=dict(orientation="h", y=1.1),
            margin=dict(l=0, r=0, t=40, b=0),
            height=320
        )

        # Create Box Plot (Distribution of values for the specific year)
        box_new = df_new[df_new['Year'] == year]
        box_exist = df_exist[df_exist['Year'] == year]

        fig_box = go.Figure()
        fig_box.add_trace(go.Box(y=box_new['Value'], name='New', marker_color='#2ecc71', boxpoints='all', jitter=0.3, pointpos=-1.8))
        fig_box.add_trace(go.Box(y=box_exist['Value'], name='Existing', marker_color='#e74c3c', boxpoints='all', jitter=0.3, pointpos=-1.8))

        fig_box.update_layout(
            title=f"Market Distribution ({year})",
            template='plotly_white',
            yaxis=dict(title='Change (%)'),
            margin=dict(l=0, r=0, t=40, b=0),
            height=320,
            showlegend=False
        )

        return html.Div([
            dcc.Graph(figure=fig_line, style={'height': '340px'}),
            html.Hr(className="my-1"),
            dcc.Graph(figure=fig_box, style={'height': '340px'})
        ])

if __name__ == '__main__':
    app.run(debug=True)