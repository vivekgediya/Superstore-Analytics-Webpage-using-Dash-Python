# Importing neccesory Libraries
import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
import plotly.express as px
import pandas as pd
import dash_table
import dash_bootstrap_components as dbc 

# Reading Excel File
df = pd.read_excel('Superstore.xlsx')

# External CSS
external_stylesheets = ['/assets/bWlwgp.css']

figures1 = px.bar(df, x=df['Year'], y=df['Sales'],color=df['Category'],hover_name="Order Date",barmode="group", title='Profit By Year per Category')
figures2 = px.bar(df, x=df['Year'], y=df['Profit'],color=df['Category'],hover_name="Order Date",barmode="group", title='Sales By Year per Category ')

# Constructor for main Web App
app = dash.Dash(__name__,external_stylesheets = external_stylesheets )
colors = {
    'background': '#111111',
    'text': '#7FDBFF'
}

# App Layout
app.layout = html.Div([

    html.Div(children=[
        html.H3(
            children='Dashborad To Analyze Global Superstore Data of U.S (2016-19)',
            style={
                'textAlign': 'center','color' : 'white', 'font-style':'bold', 'background-color':'black'})
    ]),
    
    html.Div([

        html.Div([
            # For Interactive Slider
            dcc.Graph(id='graph-with-slider')
        ],style={'width': '98%', 'display': 'inline-block', 'border':'1.5px solid black','padding': '10px 0px 10px 0px','margin':'10px 5px 5px 9px' }),

        html.Div([
            dcc.Slider(
                # Discount Column is Slider In Output
                id='discount-slider',
                min=df['Year'].min(), # Lower Bound
                max=df['Year'].max(), # Upper Bound
                value=df['Year'].min(), # Range starts with Minimum
                marks={str(disc): str(disc) for disc in df['Year'].unique()}, # Lambda Fun to Print each unique Discount % in Slider 
                step=None),
        ],style={'width': '98%', 'display': 'inline-block', 'border':'1.5px solid black','padding': '10px 0px 10px 0px','margin':'10px 5px 5px 9px' }),
    ],style={'width': '100%', 'display': 'inline-block', 'border':'3px solid black','margin': '0px 0px 5px 0px'}),

    html.Div([

        html.Div([

        html.Div([
            dcc.Dropdown(
                id='crossfilter-xaxis-column',
                options=[{'label': i, 'value': i} for i in df['State'].unique()],
                value='Fertility rate, total (births per woman)'
            ),
            dcc.RadioItems(
                id='crossfilter-xaxis-type',
                options=[{'label': i, 'value': i} for i in ['Linear', 'Log']],
                value='Linear',
                labelStyle={'display': 'inline-block'}
            )
        ],
        style={'width': '49%', 'display': 'inline-block','align-items': 'center'})

        ]),

        html.Div([
            # For Interactive Slider
            dcc.Graph(id='graph-bar-fig2',figure=figures2)
        ],style={'width': '98%', 'display': 'inline-block', 'border':'1.5px solid black','padding': '10px 0px 10px 0px','margin':'10px 5px 5px 9px' }),

        html.Div([
            # For Interactive Slider
            dcc.Graph(id='graph-bar-fig1',figure=figures1)
        ],style={'width': '98%', 'display': 'inline-block', 'border':'1.5px solid black','padding': '10px 0px 10px 0px','margin':'10px 5px 5px 9px' }),

        html.Div([

        ])

    ],style={'width': '100%', 'display': 'inline-block', 'border':'3px solid black'})
])


# Callback 
@app.callback(
    Output('graph-with-slider', 'figure'),
    [Input('discount-slider', 'value')])
    
def update_figure(selected_year):
    filtered_df = df[df.Year == selected_year]
    fig = px.scatter(filtered_df, y="Profit", x="Sales", 
                      color="State", hover_name="Category", 
                     log_x=True, size_max=55, title="'Sales v/s Profit' as per States")

    fig.update_layout(transition_duration=500)
   
    return fig

# For Run the Server
if __name__ == '__main__':
    app.run_server(debug=True)