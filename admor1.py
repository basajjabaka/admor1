import pandas as pd
import numpy as np
import os
import plotly.graph_objects as go
import plotly.express as px
import dash_bootstrap_components as dbc
from dash import dcc, html, Dash
from dash.dependencies import Input, Output


# Public URL with view access
url = "https://docs.google.com/spreadsheets/d/12wMFvYN3JPUnbom2vy62nczp2lVA_uPe/export?format=xlsx"
data = pd.read_excel(url, engine="openpyxl")  # Assuming newer format
# You might need to change "openpyxl" to "xlrd" for older Excel files

data['Issue'] = data['Issue'].str.replace('Closing your account', 'Closing an account')
data["Timely response?"] = data["Timely response?"].fillna("No")

df = data[['Complaint ID', 'Submitted via', 'Date submitted', 'Date received',
       'State', 'Product',  'Issue',
        'Company response to consumer',
       'Timely response?']]

card_complaints = dbc.Card(
    dbc.CardBody(
        [
            html.H1([html.I(className="bi bi-envelope")], className="text-nowrap"),
            html.H1("62,516"),
            html.H3(["Total Complaints"], className="text-nowrap")
        ], className="border-start border-success border-5"
    ),
    className="text-center m-4"
)


card_avgDaily = dbc.Card(
    dbc.CardBody(
        [
            html.H1([html.I(className="bi bi-bookmark-dash")], className="text-nowrap"),
            html.H1("27"),
            html.H3(["Average Daily Complaints"], className="text-nowrap")
        ], className="border-start border-danger border-5"
    ),
    className="text-center m-4",
)


card_highestYr = dbc.Card(
    dbc.CardBody(
        [
            html.H1([html.I(className="bi bi-calendar2-day")], className="text-nowrap"),
            html.H1("12953"),
            html.H3(["Highest Complaints In A Year (2022)"], className="text-nowrap")
            
        ], className="border-start border-success border-5"
    ),
    className="text-center m-4",
)



app1 = Dash(__name__, external_stylesheets=[
                "https://codepen.io/chriddyp/pen/bWLwgP.css",
                dbc.themes.SPACELAB, 
                dbc.icons.BOOTSTRAP, 
                ])

states = ['NY', 'FL', 'CA', 'VA', 'TX', 'KS', 'GA', 'CT', 'OH', 'NJ', 'IL', 'MI', 'NC', 'PA', 'WA', 
          'IN', 'MA', 'MD', 'NV', 'TN', 'AZ', 'MO', 'DC', 'ID', 'MS', 'CO', 'OR', 'MN', 'KY', 'AR', 
          'NH', 'NM', 'UT', 'SC', 'AL', 'DE', 'OK', 'LA', 'RI', 'WI', 'IA', 'ME', 'WV', 'VT', 'NE', 
          'SD', 'HI', 'AK', 'MT', 'ND', 'WY']

app1.layout = html.Div([
    html.H1("Consumer Complaints Analysis Dashboard", style={'textAlign': 'center', 'color': '#1F77B4'}),

    dbc.Container(
        dbc.Row(
            [dbc.Col(card_complaints, width=4), dbc.Col(card_avgDaily, width=4), dbc.Col(card_highestYr, width=4)],
        ), fluid=True
),
     html.Div([
        dcc.Dropdown(
        id='dropdown1',
        options=[{'label': str(state), 'value': state} for state in states],
        placeholder="Select State", value='NY'
    ),
        # Graph component
        dcc.Graph(id='bar1')
], style={'width': '33%', 'display': 'inline-block', 'padding': '10px'}),

    html.Div([
        dcc.Graph(
            id='bar3')

    ], style={'width': '33%', 'display': 'inline-block', 'padding': '10px'}),

     html.Div([
        dcc.Dropdown(
        id='dropdown2',
        options=[{'label': str(state), 'value': state} for state in states],
        placeholder="Select State", value='NY'
    ),
        # Graph component
        dcc.Graph(id='bar2')
], style={'width': '33%', 'display': 'inline-block', 'padding': '10px'}),

    html.Div([
        dcc.Graph(id="line1", style={"width":"49%", "display":"inline-block"}),
        dcc.Graph(id="dist1", style={"width":"49%", "display":"inline-block"}),
    ],  style={"display":"flex", "width":"100%"}),

    html.Div([
        dcc.Graph(id="line2",),
        
    ], style={"width": "100%", "display": "block", "padding": "10px"}),

])


@app1.callback(
    Output('bar1', 'figure'),
    [Input('dropdown1', 'value')]
)
def update_bar1(state):
    
    filtered_df = df[df['State'] == state]

    # Get the top 5 issues
    top5 = filtered_df['Issue'].value_counts().nlargest(5)
    
    fig = px.bar(top5, x = top5.index, y = top5.values,
                 labels = {'x' : 'Issues', 'y': 'Frequency'},
                 title = f'Top 5 Issues in {state}')

    fig.update_xaxes(title_text="Issues")
    return fig


@app1.callback(
    Output('bar2', 'figure'),
    [Input('dropdown2', 'value')]
)
def update_bar2(state):
    
    filtered_df = df[df['State'] == state]

    # Get the top 5 issues
    top5 = filtered_df['Product'].value_counts().nlargest(5)
    
    fig = px.bar(top5, x = top5.index, y = top5.values,
                 labels = {'x' : 'Products', 'y': 'Frequency'},
                 title = f'Top 5 Products in {state}')

    fig.update_xaxes(title_text="Products")
    return fig


@app1.callback(
    Output("bar3", "figure"),
    [Input("bar3", "id")]
)
def update_bar3(_):
    top_states = df['State'].value_counts().nlargest(5)
    fig = px.bar(
        x=top_states.index, 
        y=top_states.values,
        labels={'x': 'State', 'y': 'Complaints Frequency'},
        title='Top 5 States with Highest Complaints'
    )
    fig.update_xaxes(title_text="State")
    return fig


@app1.callback(
    Output("line1", "figure"),
    [Input("line1", "id")]
)
def update_line1(_):

    filtered_data = df[(df['Date submitted'].dt.year >= 2017) & (df['Date submitted'].dt.year <= 2022)]
    
    yearly_counts = filtered_data['Date submitted'].dt.year.value_counts().sort_index()

    fig = px.line(
        x=yearly_counts.index, 
        y=yearly_counts.values,
        labels={'x': 'Year', 'y': 'Complaint Frequency'},
        title='Complaints Frequency for all years'
    )
    fig.update_xaxes(title_text="Years")
    return fig

@app1.callback(
    Output("dist1", "figure"),
    [Input("dist1", "id")]
)
def update_dist1(_):
    
    submission_counts = df['Submitted via'].value_counts()

    fig = px.pie(
        submission_counts,
        names=submission_counts.index,
        values=submission_counts.values,
        hole=0.5,
        labels={'names': 'Submission Method', 'values': 'Count'},
        title="Complaints Submission Methods"
    )
    
    fig.update_traces(textinfo="percent+label")
    return fig

@app1.callback(
    Output("line2", "figure"),
    [Input("line2", "id")]
)
def update_seasonality(_):

    monthly_counts = df['Date submitted'].dt.month.value_counts().sort_index()
    
    fig = px.line(
        x=monthly_counts.index, 
        y=monthly_counts.values,
        labels={'x': 'Month', 'y': 'Complaints Frequency'},
        title='Seasonality of Complaints by Month'
    )
    
    # Set month labels on x-axis
    fig.update_xaxes(
        title_text="Month",
        tickvals=list(range(1, 13)),
        ticktext=["January", "February", "March", "April", "May", "June", 
                  "July", "August", "September", "October", "November", "December"]
    )
    
    fig.update_yaxes(title_text="Complaints Frequency")
    return fig



app1.run_server(debug=True, mode='inline', port=5450)
