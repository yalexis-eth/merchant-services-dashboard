import os
import dash
from dash import dcc, html, dash_table
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash.dependencies import Input, Output, State
import base64
import io
import re
from dateutil.parser import parse
import dash_bootstrap_components as dbc
from dash.dash_table.Format import Format, Scheme, Group

# Initialize app with a professional theme
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)

# Define volume columns for Total Volume calculation
volume_columns = [
    'V/MC/Discover Vol', 'AMEX Vol', 'Wex Voyager Volume', 'EBT Vol',
    'PIN DEB Vol', 'VISA MC REF vol', 'MCP Volume'
]

# Define base columns that are always available
base_mid_columns = [
    {'name': 'MID', 'id': 'MID', 'type': 'text'},
    {'name': 'DBA Name', 'id': 'DBA Name', 'type': 'text'},
    {'name': 'Total Volume', 'id': 'Total Volume', 'type': 'numeric', 
     'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
    {'name': 'Agent Net', 'id': 'Agent Net', 'type': 'numeric', 
     'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
    {'name': 'Gross Margin %', 'id': 'Gross Margin %', 'type': 'numeric', 
     'format': Format(precision=2, scheme=Scheme.fixed, symbol_suffix='%')},
    {'name': 'V/MC/Discover Vol', 'id': 'V/MC/Discover Vol', 'type': 'numeric',
     'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
    {'name': 'AMEX Vol', 'id': 'AMEX Vol', 'type': 'numeric',
     'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
    {'name': 'Wex Voyager Volume', 'id': 'Wex Voyager Volume', 'type': 'numeric',
     'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
    {'name': 'EBT Vol', 'id': 'EBT Vol', 'type': 'numeric',
     'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
    {'name': 'PIN DEB Vol', 'id': 'PIN DEB Vol', 'type': 'numeric',
     'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
    {'name': 'VISA MC REF vol', 'id': 'VISA MC REF vol', 'type': 'numeric',
     'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
    {'name': 'MCP Volume', 'id': 'MCP Volume', 'type': 'numeric',
     'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)}
]

# Default visible columns
default_visible_columns = ['MID', 'DBA Name', 'Total Volume', 'Agent Net', 'Gross Margin %']

# Custom styles
CARD_STYLE = {
    'box-shadow': '0 4px 6px 0 rgba(0, 0, 0, 0.1)',
    'border-radius': '10px',
    'padding': '20px',
    'margin-bottom': '20px',
    'background-color': 'white'
}

# Layout
app.layout = html.Div([
    # Header
    dbc.NavbarSimple(
        brand="B2B Merchant Services .xls App",
        brand_href="#",
        color="primary",
        dark=True,
        className="mb-4"
    ),
    
    dbc.Container([
        # Upload section
        dbc.Card([
            dbc.CardBody([
                html.H4("Data Upload", className="card-title mb-3"),
                dbc.Row([
                    dbc.Col([
                        dcc.Upload(
                            id='upload-data',
                            children=dbc.Button(
                                [html.I(className="fas fa-upload me-2"), "Upload .xls Files"],
                                color="primary",
                                size="lg"
                            ),
                            multiple=True
                        ),
                    ], width=6),
                    dbc.Col([
                        dbc.Button(
                            [html.I(className="fas fa-trash me-2"), "Clear All Files"],
                            id='clear-button',
                            color="danger",
                            size="lg"
                        ),
                    ], width=6),
                ]),
                html.Div(id='file-list', className='mt-3'),
            ])
        ], className="mb-4"),
        
        dcc.Store(id='stored-data'),
        dcc.Store(id='available-columns-store'),
        
        # KPI Cards Summary
        html.Div(id='kpi-cards', className='mb-4'),
        
        # Summary section with enhanced table
        html.Div(id='summary-section', className='mb-4'),
        
        # Charts section
        html.Div(id='charts-section', className='mb-4'),
        
        # Store for filtered data
        dcc.Store(id='filtered-mid-data'),
        
        # Individual MID margins section with column selector
        dbc.Card([
            dbc.CardBody([
                html.H4('Individual MID Analysis', className="card-title mb-3"),
                
                # Filters Row
                dbc.Row([
                    dbc.Col([
                        dcc.Dropdown(
                            id='month-dropdown',
                            placeholder='Select Month',
                            className='mb-2'
                        ),
                    ], width=6),
                    dbc.Col([
                        dcc.Dropdown(
                            id='filter-dropdown',
                            options=[
                                {'label': 'All Margins', 'value': 'all'},
                                {'label': 'Positive Margins Only', 'value': 'positive'},
                                {'label': 'Negative Margins Only', 'value': 'negative'},
                                {'label': 'High Margins (>5%)', 'value': 'high'},
                                {'label': 'Low Margins (<1%)', 'value': 'low'},
                                {'label': 'Improving MIDs (↑)', 'value': 'improving'},
                                {'label': 'Declining MIDs (↓)', 'value': 'declining'}
                            ],
                            placeholder='Select Filter',
                            value='all',
                            className='mb-2'
                        ),
                    ], width=6),
                ]),
                
                # Column Selector
                dbc.Card([
                    dbc.CardBody([
                        html.H6("Select Columns to Display", className="mb-3"),
                        html.Div(id='column-selector-container'),
                        dbc.Row([
                            dbc.Col([
                                dbc.ButtonGroup([
                                    dbc.Button("Select All", id="select-all-btn", color="secondary", size="sm"),
                                    dbc.Button("Clear All", id="clear-all-btn", color="secondary", size="sm"),
                                    dbc.Button("Reset Default", id="reset-default-btn", color="secondary", size="sm"),
                                ], className="mt-2")
                            ], width=12),
                        ])
                    ])
                ], color="light", className="mb-3"),
                
                # Hidden dummy components to prevent callback errors
                html.Div([
                    dbc.Checklist(id='column-selector', options=[], value=[], style={'display': 'none'}),
                    dbc.Checklist(id='volume-columns-selector', options=[], value=[], style={'display': 'none'}),
                    dbc.Checklist(id='margin-columns-selector', options=[], value=[], style={'display': 'none'}),
                    dbc.Checklist(id='change-columns-selector', options=[], value=[], style={'display': 'none'}),
                ]),
                
                # Table container
                html.Div(id='mid-table-container'),
                
                # Export button
                dbc.Button(
                    [html.I(className="fas fa-download me-2"), "Export to CSV"],
                    id='export-button',
                    color="success",
                    className='mt-3'
                ),
                dcc.Download(id='download-csv')
            ])
        ]),
    ], fluid=True)
], style={'background-color': '#f5f5f5', 'min-height': '100vh'})

# Helper functions
def extract_month_year(filename):
    match = re.search(r'- (\w+ \d{4})\.xls', filename)
    return parse(match.group(1)) if match else None

def clean_data(df):
    """Clean the dataframe by converting columns to numeric and filtering out total rows."""
    for col in volume_columns + ['Agent Net']:
        if col in df.columns:
            df[col] = df[col].replace('[\$,]', '', regex=True).astype(float)
    df['Total Volume'] = df[volume_columns].sum(axis=1)
    df['Gross Margin %'] = df.apply(
        lambda row: (row['Agent Net'] / row['Total Volume']) * 100 if row['Total Volume'] > 0 else float('nan'),
        axis=1
    )
    # Remove rows where 'MID' is NaN
    df = df[df['MID'].notna()]
    # Convert 'MID' to string for string operations
    df['MID'] = df['MID'].astype(str)
    # Filter out rows where 'MID' contains 'total' (case-insensitive)
    df = df[~df['MID'].str.contains('total', case=False)]
    # Remove duplicates based on 'MID'
    df = df.drop_duplicates(subset=['MID'], keep='first')
    return df

def create_kpi_card(title, value, change=None, icon="fas fa-chart-line", format_currency=False):
    """Create a KPI card with optional change indicator"""
    if format_currency:
        value_display = f"${value:,.2f}"
        change_display = f"${change:,.2f}" if change is not None else None
    else:
        value_display = f"{value:,}"
        change_display = f"{change:+,.0f}" if change is not None else None
    
    # Determine color based on change
    if change is not None:
        if change > 0:
            change_color = "success"
            arrow_icon = "fas fa-arrow-up"
        elif change < 0:
            change_color = "danger"
            arrow_icon = "fas fa-arrow-down"
        else:
            change_color = "secondary"
            arrow_icon = "fas fa-minus"
    
    card_content = [
        dbc.Row([
            dbc.Col([
                html.I(className=f"{icon} fa-2x text-primary")
            ], width=3),
            dbc.Col([
                html.H6(title, className="text-muted mb-1"),
                html.H3(value_display, className="mb-0"),
                html.P([
                    html.I(className=f"{arrow_icon} me-1") if change is not None else "",
                    change_display
                ], className=f"text-{change_color} mb-0 small") if change is not None else None
            ], width=9)
        ])
    ]
    
    return dbc.Card(card_content, style=CARD_STYLE)

# Callback to update available columns based on uploaded data
@app.callback(
    Output('available-columns-store', 'data'),
    Input('stored-data', 'data')
)
def update_available_columns(data):
    if not data:
        return base_mid_columns
    
    # Get all months sorted
    sorted_months = sorted(data.keys(), key=lambda x: parse(x))
    
    # Create dynamic columns for each month's margin
    all_columns = base_mid_columns.copy()
    
    # Add columns for each month's margin
    for month in sorted_months:
        month_column = {
            'name': f'{month} Margin %',
            'id': f'{month} Margin %',
            'type': 'numeric',
            'format': Format(precision=2, scheme=Scheme.fixed, symbol_suffix='%')
        }
        all_columns.append(month_column)
    
    # Add margin change columns
    change_columns = []
    for i in range(1, len(sorted_months)):
        prev_month = sorted_months[i-1]
        curr_month = sorted_months[i]
        change_column = {
            'name': f'Change {prev_month} → {curr_month}',
            'id': f'Change_{prev_month}_{curr_month}',
            'type': 'numeric',
            'format': Format(precision=2, scheme=Scheme.fixed, symbol_suffix='pp')
        }
        change_columns.append(change_column)
    
    all_columns.extend(change_columns)
    
    return all_columns

# Callback to update column selector
@app.callback(
    [Output('column-selector', 'options'),
     Output('column-selector', 'value'),
     Output('volume-columns-selector', 'options'),
     Output('volume-columns-selector', 'value'),
     Output('margin-columns-selector', 'options'),
     Output('margin-columns-selector', 'value'),
     Output('change-columns-selector', 'options'),
     Output('change-columns-selector', 'value'),
     Output('column-selector-container', 'children')],
    [Input('available-columns-store', 'data'), Input('stored-data', 'data')]
)
def update_column_selector(available_columns, data):
    if not available_columns or not data:
        return [], [], [], [], [], [], [], [], html.Div("No data available")
    
    # Get current month's columns as default
    sorted_months = sorted(data.keys(), key=lambda x: parse(x)) if data else []
    current_month = sorted_months[-1] if sorted_months else None
    
    # Default selections include base columns plus current month margin
    default_selections = default_visible_columns.copy()
    if current_month:
        default_selections.append(f'{current_month} Margin %')
    
    # Group columns by type
    basic_columns = []
    volume_columns_group = []
    margin_columns = []
    change_columns = []
    
    for col in available_columns:
        if col['id'] in ['MID', 'DBA Name', 'Total Volume', 'Agent Net', 'Gross Margin %']:
            basic_columns.append({'label': col['name'], 'value': col['id']})
        elif 'Vol' in col['id'] or 'Volume' in col['id'] and col['id'] != 'Total Volume':
            volume_columns_group.append({'label': col['name'], 'value': col['id']})
        elif 'Margin %' in col['name'] and 'Change' not in col['name']:
            margin_columns.append({'label': col['name'], 'value': col['id']})
        elif 'Change' in col['id']:
            change_columns.append({'label': col['name'], 'value': col['id']})
    
    # Prepare values
    basic_values = [opt['value'] for opt in basic_columns if opt['value'] in default_selections]
    volume_values = []
    margin_values = [opt['value'] for opt in margin_columns if opt['value'] in default_selections]
    change_values = []
    
    # Create the visual container
    container = html.Div([
        # Basic Information
        html.Div([
            html.H6("Basic Information", className="text-muted mb-2"),
            dbc.Checklist(
                id='column-selector-display',
                options=basic_columns,
                value=basic_values,
                inline=True,
                className='mb-3',
                style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '15px'}
            )
        ]),
        
        # Volume Breakdown
        html.Div([
            html.H6("Volume Breakdown", className="text-muted mb-2"),
            dbc.Checklist(
                id='volume-columns-selector-display',
                options=volume_columns_group,
                value=volume_values,
                inline=True,
                className='mb-3',
                style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '15px'}
            )
        ]) if volume_columns_group else None,
        
        # Monthly Margins
        html.Div([
            html.H6("Monthly Margins", className="text-muted mb-2"),
            dbc.Checklist(
                id='margin-columns-selector-display',
                options=margin_columns,
                value=margin_values,
                inline=True,
                className='mb-3',
                style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '15px'}
            )
        ]) if margin_columns else None,
        
        # Month-to-Month Changes
        html.Div([
            html.H6("Month-to-Month Changes", className="text-muted mb-2"),
            dbc.Checklist(
                id='change-columns-selector-display',
                options=change_columns,
                value=change_values,
                inline=True,
                className='mb-3',
                style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '15px'}
            )
        ]) if change_columns else None,
    ])
    
    return (basic_columns, basic_values, 
            volume_columns_group, volume_values,
            margin_columns, margin_values,
            change_columns, change_values,
            container)

# Add client-side callback to sync display checklists with hidden ones
app.clientside_callback(
    """
    function(basic, volume, margin, change) {
        return [basic || [], volume || [], margin || [], change || []];
    }
    """,
    [Output('column-selector', 'value', allow_duplicate=True),
     Output('volume-columns-selector', 'value', allow_duplicate=True),
     Output('margin-columns-selector', 'value', allow_duplicate=True),
     Output('change-columns-selector', 'value', allow_duplicate=True)],
    [Input('column-selector-display', 'value'),
     Input('volume-columns-selector-display', 'value'),
     Input('margin-columns-selector-display', 'value'),
     Input('change-columns-selector-display', 'value')],
    prevent_initial_call=True
)

# Update button callbacks to work with all selectors
@app.callback(
    [Output('column-selector-display', 'value'),
     Output('volume-columns-selector-display', 'value'),
     Output('margin-columns-selector-display', 'value'),
     Output('change-columns-selector-display', 'value')],
    [Input('select-all-btn', 'n_clicks'),
     Input('clear-all-btn', 'n_clicks'),
     Input('reset-default-btn', 'n_clicks')],
    [State('column-selector', 'options'),
     State('volume-columns-selector', 'options'),
     State('margin-columns-selector', 'options'),
     State('change-columns-selector', 'options')],
    prevent_initial_call=True
)
def update_all_column_selections(select_all, clear_all, reset_default, basic_opts, vol_opts, margin_opts, change_opts):
    ctx = dash.callback_context
    if not ctx.triggered:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update
    
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'select-all-btn':
        return ([opt['value'] for opt in (basic_opts or [])],
                [opt['value'] for opt in (vol_opts or [])],
                [opt['value'] for opt in (margin_opts or [])],
                [opt['value'] for opt in (change_opts or [])])
    elif button_id == 'clear-all-btn':
        return [], [], [], []
    elif button_id == 'reset-default-btn':
        return ([opt['value'] for opt in (basic_opts or []) if opt['value'] in default_visible_columns],
                [],
                [opt['value'] for opt in (margin_opts or []) if opt['value'] in default_visible_columns],
                [])
    
    return dash.no_update, dash.no_update, dash.no_update, dash.no_update

# Main data upload callback
@app.callback(
    [Output('stored-data', 'data'), Output('file-list', 'children'), Output('month-dropdown', 'options')],
    [Input('upload-data', 'contents'), Input('clear-button', 'n_clicks')],
    [State('upload-data', 'filename'), State('stored-data', 'data')]
)
def update_data(contents, clear_clicks, filenames, existing_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return {}, [], []
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    data = existing_data or {}
    if trigger_id == 'clear-button':
        return {}, dbc.Alert("All files cleared.", color="info"), []
    if contents:
        for content, filename in zip(contents, filenames):
            month_year = extract_month_year(filename)
            if month_year:
                content_type, content_string = content.split(',')
                decoded = base64.b64decode(content_string)
                # Skip the last row (likely a total row) when reading Excel
                df = pd.read_excel(io.BytesIO(decoded), sheet_name='PPI', skipfooter=1)
                df = clean_data(df)
                data[month_year.strftime('%B %Y')] = df.to_dict('records')
    
    if data:
        file_badges = [
            dbc.Badge(f, color="primary", className="me-2")
            for f in sorted(data.keys(), key=lambda x: parse(x))
        ]
        file_display = html.Div([
            html.H6("Uploaded Files:", className="mb-2"),
            html.Div(file_badges)
        ])
    else:
        file_display = dbc.Alert("No files uploaded yet.", color="warning")
    
    month_options = [{'label': m, 'value': m} for m in sorted(data.keys(), key=lambda x: parse(x))]
    return data, file_display, month_options

# Dashboard update callback
@app.callback(
    [Output('kpi-cards', 'children'), Output('summary-section', 'children'), Output('charts-section', 'children')],
    Input('stored-data', 'data')
)
def update_dashboard(data):
    if not data:
        return [], dbc.Alert('Please upload files to view analytics.', color='info'), []
    
    summary = []
    for month, records in sorted(data.items(), key=lambda x: parse(x[0])):
        df = pd.DataFrame(records)
        total_mids = len(df)
        processing_mids = (df['Total Volume'] > 0).sum()
        positive_net_mids = (df['Agent Net'] > 0).sum()
        total_profit = df['Agent Net'].sum()
        mid_volume = df['Total Volume'].sum()
        summary.append({
            'MONTH': month,
            'TOTAL MIDS': total_mids,
            'PROCESSING MIDS': processing_mids,
            'POSITIVE NET MIDS': positive_net_mids,
            'TOTAL PROFIT': total_profit,
            'MID VOLUME': mid_volume
        })
    summary_df = pd.DataFrame(summary)
    
    # Calculate changes
    for col in ['TOTAL MIDS', 'PROCESSING MIDS', 'POSITIVE NET MIDS']:
        summary_df[f'{col} CHANGE'] = summary_df[col].diff().fillna(0)
    for col in ['TOTAL PROFIT', 'MID VOLUME']:
        summary_df[f'{col} CHANGE'] = summary_df[col].diff().fillna(0)
    
    # Get latest month data for KPI cards
    latest = summary_df.iloc[-1]
    latest_change = summary_df.iloc[-1] if len(summary_df) > 1 else None
    
    # Create KPI cards
    kpi_cards = dbc.Row([
        dbc.Col(create_kpi_card(
            "Total Profit", 
            latest['TOTAL PROFIT'],
            latest['TOTAL PROFIT CHANGE'] if latest_change is not None else None,
            "fas fa-dollar-sign",
            True
        ), width=3),
        dbc.Col(create_kpi_card(
            "MID Volume",
            latest['MID VOLUME'],
            latest['MID VOLUME CHANGE'] if latest_change is not None else None,
            "fas fa-chart-line",
            True
        ), width=3),
        dbc.Col(create_kpi_card(
            "Total MIDs",
            latest['TOTAL MIDS'],
            latest['TOTAL MIDS CHANGE'] if latest_change is not None else None,
            "fas fa-credit-card"
        ), width=3),
        dbc.Col(create_kpi_card(
            "Processing MIDs",
            latest['PROCESSING MIDS'],
            latest['PROCESSING MIDS CHANGE'] if latest_change is not None else None,
            "fas fa-check-circle"
        ), width=3),
    ])
    
    # Create enhanced summary table with conditional formatting
    def style_data_conditional():
        conditions = []
        # Profit column styling
        conditions.extend([
            {
                'if': {
                    'filter_query': '{TOTAL PROFIT} > 0',
                    'column_id': 'TOTAL PROFIT'
                },
                'backgroundColor': '#d4edda',
                'color': '#155724',
            },
            {
                'if': {
                    'filter_query': '{TOTAL PROFIT} < 0',
                    'column_id': 'TOTAL PROFIT'
                },
                'backgroundColor': '#f8d7da',
                'color': '#721c24',
            }
        ])
        # Change columns styling
        for col in ['TOTAL PROFIT CHANGE', 'MID VOLUME CHANGE', 'TOTAL MIDS CHANGE', 
                    'PROCESSING MIDS CHANGE', 'POSITIVE NET MIDS CHANGE']:
            conditions.extend([
                {
                    'if': {
                        'filter_query': f'{{{col}}} > 0',
                        'column_id': col
                    },
                    'backgroundColor': '#d1ecf1',
                    'color': '#0c5460',
                },
                {
                    'if': {
                        'filter_query': f'{{{col}}} < 0',
                        'column_id': col
                    },
                    'backgroundColor': '#f8d7da',
                    'color': '#721c24',
                }
            ])
        return conditions
    
    # Define table columns
    columns = [
        {'name': 'Month', 'id': 'MONTH', 'type': 'text'},
        {'name': 'Total MIDs', 'id': 'TOTAL MIDS', 'type': 'numeric'},
        {'name': 'Processing', 'id': 'PROCESSING MIDS', 'type': 'numeric'},
        {'name': 'Positive Net', 'id': 'POSITIVE NET MIDS', 'type': 'numeric'},
        {'name': 'Total Profit', 'id': 'TOTAL PROFIT', 'type': 'numeric', 
         'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
        {'name': 'Volume', 'id': 'MID VOLUME', 'type': 'numeric', 
         'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
        {'name': 'MIDs Δ', 'id': 'TOTAL MIDS CHANGE', 'type': 'numeric'},
        {'name': 'Processing Δ', 'id': 'PROCESSING MIDS CHANGE', 'type': 'numeric'},
        {'name': 'Positive Δ', 'id': 'POSITIVE NET MIDS CHANGE', 'type': 'numeric'},
        {'name': 'Profit Δ', 'id': 'TOTAL PROFIT CHANGE', 'type': 'numeric', 
         'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)},
        {'name': 'Volume Δ', 'id': 'MID VOLUME CHANGE', 'type': 'numeric', 
         'format': Format(symbol_prefix="$", precision=2, scheme=Scheme.fixed, group=Group.yes)}
    ]
    
    summary_table = dbc.Card([
        dbc.CardBody([
            html.H4("Monthly Summary Metrics", className="card-title mb-3"),
            dash_table.DataTable(
                columns=columns,
                data=summary_df.to_dict('records'),
                style_cell={
                    'textAlign': 'center',
                    'padding': '10px',
                    'fontFamily': 'Arial'
                },
                style_header={
                    'backgroundColor': '#007bff',
                    'color': 'white',
                    'fontWeight': 'bold'
                },
                style_data_conditional=style_data_conditional(),
                style_table={'overflowX': 'auto'}
            )
        ])
    ])
    
    # Create enhanced charts
    # 1. Combined Profit and Volume Chart
    fig_combined = go.Figure()
    fig_combined.add_trace(go.Bar(
        x=summary_df['MONTH'],
        y=summary_df['TOTAL PROFIT'],
        name='Total Profit',
        marker_color=['#28a745' if x > 0 else '#dc3545' for x in summary_df['TOTAL PROFIT']],
        yaxis='y'
    ))
    fig_combined.add_trace(go.Scatter(
        x=summary_df['MONTH'],
        y=summary_df['MID VOLUME'],
        name='MID Volume',
        line=dict(color='#17a2b8', width=3),
        yaxis='y2'
    ))
    fig_combined.update_layout(
        title='Profit vs Volume Trend',
        xaxis_title='Month',
        yaxis=dict(title='Total Profit ($)', side='left'),
        yaxis2=dict(title='MID Volume ($)', side='right', overlaying='y'),
        hovermode='x unified',
        template='plotly_white',
        height=400
    )
    
    # 2. MIDs Performance Chart
    fig_mids = go.Figure()
    fig_mids.add_trace(go.Scatter(
        x=summary_df['MONTH'],
        y=summary_df['TOTAL MIDS'],
        name='Total MIDs',
        line=dict(color='#6c757d', width=2)
    ))
    fig_mids.add_trace(go.Scatter(
        x=summary_df['MONTH'],
        y=summary_df['PROCESSING MIDS'],
        name='Processing MIDs',
        line=dict(color='#28a745', width=2)
    ))
    fig_mids.add_trace(go.Scatter(
        x=summary_df['MONTH'],
        y=summary_df['POSITIVE NET MIDS'],
        name='Positive Net MIDs',
        line=dict(color='#ffc107', width=2)
    ))
    fig_mids.update_layout(
        title='MIDs Performance Overview',
        xaxis_title='Month',
        yaxis_title='Number of MIDs',
        hovermode='x unified',
        template='plotly_white',
        height=400
    )
    
    # 3. Profit Margin Gauge for latest month
    if len(summary_df) > 0:
        latest_margin = (latest['TOTAL PROFIT'] / latest['MID VOLUME'] * 100) if latest['MID VOLUME'] > 0 else 0
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=latest_margin,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': f"Overall Profit Margin % ({latest['MONTH']})"},
            delta={'reference': summary_df.iloc[-2]['TOTAL PROFIT'] / summary_df.iloc[-2]['MID VOLUME'] * 100 
                   if len(summary_df) > 1 and summary_df.iloc[-2]['MID VOLUME'] > 0 else 0},
            gauge={
                'axis': {'range': [None, 10]},
                'bar': {'color': "#28a745" if latest_margin > 0 else "#dc3545"},
                'steps': [
                    {'range': [0, 1], 'color': "#f8d7da"},
                    {'range': [1, 3], 'color': "#fff3cd"},
                    {'range': [3, 5], 'color': "#d1ecf1"},
                    {'range': [5, 10], 'color': "#d4edda"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 0
                }
            }
        ))
        fig_gauge.update_layout(height=300)
    
    charts = dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardBody(dcc.Graph(figure=fig_combined))
            ])
        ], width=12, className="mb-3"),
        dbc.Col([
            dbc.Card([
                dbc.CardBody(dcc.Graph(figure=fig_mids))
            ])
        ], width=8),
        dbc.Col([
            dbc.Card([
                dbc.CardBody(dcc.Graph(figure=fig_gauge))
            ])
        ], width=4),
    ])
    
    return kpi_cards, summary_table, charts

# MID table update callback with column selection
@app.callback(
    [Output('mid-table-container', 'children'), Output('filtered-mid-data', 'data')],
    [Input('month-dropdown', 'value'), 
     Input('filter-dropdown', 'value'),
     Input('column-selector', 'value'),
     Input('volume-columns-selector', 'value'),
     Input('margin-columns-selector', 'value'),
     Input('change-columns-selector', 'value')],
    State('stored-data', 'data')
)
def update_mid_table(selected_month, filter_type, basic_cols, vol_cols, margin_cols, change_cols, data):
    if not data or not selected_month:
        return dbc.Alert("Please select a month to view MID details.", color="info"), []
    
    # Combine all selected columns
    selected_columns = (basic_cols or []) + (vol_cols or []) + (margin_cols or []) + (change_cols or [])
    
    if not selected_columns:
        return dbc.Alert("Please select at least one column to display.", color="warning"), []
    
    # Get all months sorted
    sorted_months = sorted(data.keys(), key=lambda x: parse(x))
    
    # Start with the selected month's data
    df = pd.DataFrame(data[selected_month])
    
    # Add margins from all months
    for month in sorted_months:
        if month != selected_month:
            month_df = pd.DataFrame(data[month])
            df = df.merge(
                month_df[['MID', 'Gross Margin %']].rename(columns={'Gross Margin %': f'{month} Margin %'}),
                on='MID',
                how='left'
            )
    
    # Rename current month's margin column
    df = df.rename(columns={'Gross Margin %': f'{selected_month} Margin %'})
    df['Gross Margin %'] = df[f'{selected_month} Margin %']  # Keep for filtering
    
    # Calculate month-to-month changes
    for i in range(1, len(sorted_months)):
        prev_month = sorted_months[i-1]
        curr_month = sorted_months[i]
        change_col_name = f'Change_{prev_month}_{curr_month}'
        
        if f'{prev_month} Margin %' in df.columns and f'{curr_month} Margin %' in df.columns:
            df[change_col_name] = df[f'{curr_month} Margin %'] - df[f'{prev_month} Margin %']
    
    # Apply filters based on current month's margin
    if filter_type == 'positive':
        df = df[df['Gross Margin %'] > 0]
    elif filter_type == 'negative':
        df = df[df['Gross Margin %'] < 0]
    elif filter_type == 'high':
        df = df[df['Gross Margin %'] > 5]
    elif filter_type == 'low':
        df = df[df['Gross Margin %'] < 1]
    elif filter_type == 'improving':
        # Find the change column that ends with current month
        for col in df.columns:
            if col.startswith('Change_') and col.endswith(f'_{selected_month}'):
                df = df[df[col] > 0]
                break
    elif filter_type == 'declining':
        # Find the change column that ends with current month
        for col in df.columns:
            if col.startswith('Change_') and col.endswith(f'_{selected_month}'):
                df = df[df[col] < 0]
                break
    
    # Sort by volume descending
    df = df.sort_values('Total Volume', ascending=False)
    
    # Store filtered data for export (include all columns)
    export_columns = ['MID', 'DBA Name', 'Total Volume', 'Agent Net'] + \
                    [col for col in df.columns if 'Margin %' in col or col.startswith('Change_')] + \
                    volume_columns
    export_columns = [col for col in export_columns if col in df.columns]
    filtered_data = df[export_columns].to_dict('records')
    
    # Create summary stats
    total_records = len(df)
    avg_margin = df['Gross Margin %'].mean()
    total_volume = df['Total Volume'].sum()
    
    # Count MIDs with data from previous months
    mids_with_history = 0
    for month in sorted_months:
        if month != selected_month and f'{month} Margin %' in df.columns:
            mids_with_history = df[f'{month} Margin %'].notna().sum()
            break
    
    stats = dbc.Alert([
        html.H6("Quick Stats", className="alert-heading"),
        html.P([
            f"Total Records: {total_records} | ",
            f"Avg Current Margin: {avg_margin:.2f}% | ",
            f"Total Volume: ${total_volume:,.2f}"
        ], className="mb-1"),
        html.P([
            f"MIDs with Historical Data: {mids_with_history} | ",
            f"Months Available: {len(sorted_months)}"
        ], className="mb-0 text-muted small")
    ], color="light")
    
    # Build column definitions for visible columns
    all_available_columns = []
    
    # Add base columns
    for col in base_mid_columns:
        if col['id'] in selected_columns:
            all_available_columns.append(col)
    
    # Add month margin columns
    for month in sorted_months:
        col_id = f'{month} Margin %'
        if col_id in selected_columns and col_id in df.columns:
            all_available_columns.append({
                'name': col_id,
                'id': col_id,
                'type': 'numeric',
                'format': Format(precision=2, scheme=Scheme.fixed, symbol_suffix='%')
            })
    
    # Add change columns
    for i in range(1, len(sorted_months)):
        prev_month = sorted_months[i-1]
        curr_month = sorted_months[i]
        col_id = f'Change_{prev_month}_{curr_month}'
        if col_id in selected_columns and col_id in df.columns:
            all_available_columns.append({
                'name': f'Change {prev_month} → {curr_month}',
                'id': col_id,
                'type': 'numeric',
                'format': Format(precision=2, scheme=Scheme.fixed, symbol_suffix='pp')
            })
    
    # Style conditions for the table
    style_data_conditional = []
    
    # Style all margin columns
    for month in sorted_months:
        col_id = f'{month} Margin %'
        if col_id in df.columns:
            style_data_conditional.extend([
                {
                    'if': {
                        'filter_query': f'{{{col_id}}} > 5',
                        'column_id': col_id
                    },
                    'backgroundColor': '#28a745',
                    'color': 'white',
                },
                {
                    'if': {
                        'filter_query': f'{{{col_id}}} > 0 && {{{col_id}}} <= 5',
                        'column_id': col_id
                    },
                    'backgroundColor': '#d4edda',
                    'color': '#155724',
                },
                {
                    'if': {
                        'filter_query': f'{{{col_id}}} < 0',
                        'column_id': col_id
                    },
                    'backgroundColor': '#dc3545',
                    'color': 'white',
                }
            ])
    
    # Style change columns
    for col in df.columns:
        if col.startswith('Change_'):
            style_data_conditional.extend([
                {
                    'if': {
                        'filter_query': f'{{{col}}} > 0',
                        'column_id': col
                    },
                    'backgroundColor': '#d1ecf1',
                    'color': '#0c5460',
                },
                {
                    'if': {
                        'filter_query': f'{{{col}}} < 0',
                        'column_id': col
                    },
                    'backgroundColor': '#f8d7da',
                    'color': '#721c24',
                }
            ])
    
    # Prepare data for visible columns only
    table_data = []
    for record in df.to_dict('records'):
        row = {}
        for col in selected_columns:
            if col in record:
                row[col] = record[col]
        table_data.append(row)
    
    table = dash_table.DataTable(
        id='mid-table',
        columns=all_available_columns,
        data=table_data,
        filter_action='native',
        sort_action='native',
        page_size=15,
        style_cell={'textAlign': 'center', 'padding': '10px'},
        style_header={
            'backgroundColor': '#007bff',
            'color': 'white',
            'fontWeight': 'bold'
        },
        style_data_conditional=style_data_conditional,
        style_table={'overflowX': 'auto'},
        tooltip_duration=None
    )
    
    # Add note about selected columns
    selected_note = html.P([
        html.I(className="fas fa-info-circle me-2"),
        f"Displaying {len(selected_columns)} of {len(base_mid_columns) + len(sorted_months) + max(0, len(sorted_months)-1)} available columns. ",
        f"Current month: {selected_month}"
    ], className="text-muted small mt-2")
    
    return html.Div([stats, table, selected_note]), filtered_data

@app.callback(
    Output('download-csv', 'data'),
    Input('export-button', 'n_clicks'),
    [State('filtered-mid-data', 'data'), State('month-dropdown', 'value')]
)
def export_csv(n_clicks, filtered_data, selected_month):
    if n_clicks and filtered_data:
        df = pd.DataFrame(filtered_data)
        return dcc.send_data_frame(df.to_csv, f"gross_margin_{selected_month}_comparison.csv", index=False)
    return None

if __name__ == '__main__':
    app.run(debug=True)
else:
    app.run_server(host='0.0.0.0', port=int(os.environ.get('PORT', 8050)))
