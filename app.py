from dash import Dash, dcc, html, Input, Output
import pandas as pd
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
import plotly.express as px

# pull data from Excel spreadsheet
dailydata = pd.read_excel("Region_US48.xlsx", "Published Daily Data")
hourlydata = pd.read_excel("Region_US48.xlsx", "Published Hourly Data")

# Add a date and year column to the dailydata dataframe
dailydata['date'] = pd.to_datetime(dailydata['Local date'])
dailydata['year'] = dailydata['Local date'].dt.year

# add a date and month column to the houlydata datafram
hourlydata['date'] = pd.to_datetime(hourlydata["Local date"])
hourlydata['month'] = hourlydata['date'].dt.month

# create a pivot table for a heatmap of NG (Net Generation) by hour
heatmapdata = hourlydata.pivot_table(index='Hour', columns="month", values='NG', aggfunc='mean')
heatmapdata = heatmapdata.iloc[:-1, :] # remove null values
heatmapdata.index = heatmapdata.index.astype(int)

# pull the CO2 Emissions Generated column from the hourlydata dataframe
emissiondata = hourlydata['CO2 Emissions Generated']

# store unique years for slicer
unique_years = dailydata['year'].unique()
year_options = [{'label': str(year), 'value': year} for year in unique_years]

# Create the heatmap figure
heatmap = go.Figure(data=go.Heatmap(z=heatmapdata.values.tolist(), x=list(heatmapdata.columns), y=list(heatmapdata.index), colorscale='Viridis'))
heatmap.update_layout(title="Average Hourly Net Generation Heatmap", xaxis_title="Month", yaxis_title="Hour",
                      xaxis=dict(tickmode='array', tickvals=list(heatmapdata.columns), ticktext=['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']),
                      yaxis=dict(tickmode='array', tickvals=list(heatmapdata.index), ticktext=list(heatmapdata.index)))

# dash app w/ theme
app = Dash(external_stylesheets=[dbc.themes.PULSE], suppress_callback_exceptions=True)

# sidebar layout
sidebar = html.Div(
    [
        html.H4("Select data to view:"),
        html.Hr(),
        dbc.Nav(
            [
                dbc.NavLink("Generation", href="/generation", id="generation-link", style={"background-color":"WhiteSmoke", "color": "black", "text-decoration": "none", "padding": "10px", "border": "0.5px solid black", "border-radius": "5px", "text-align": "center", "margin-bottom": "5px"}),
                dbc.NavLink("Emissions", href="/emissions", id="emissions-link", style={"background-color":"WhiteSmoke","color": "black", "text-decoration": "none", "padding": "10px", "border": "0.5px solid black", "border-radius": "5px", "text-align": "center", "margin-bottom": "5px"}),
            ],
            vertical=True,
            pills=True,
        ),
        html.Hr(),
        html.H5('Select a year to view: '),
        dcc.Dropdown(
            id='year-dropdown',
            options=year_options,
            value=unique_years[0],
            style={'color':'black'}
        ),
    ],
    style={"position": "fixed", "top": 0, "left": 0, "bottom": 0, "width": "16rem", "padding": "2rem 1rem", "background-color": "royalblue", "color":"white"},
)

#  main content layout
content = html.Div(id="page-content", style={"padding-left": "16rem"})

# app layout (combine sidebar and main content)
app.layout = html.Div([
    dcc.Location(id="url", refresh=False, pathname="/generation"),  # initialize with /generation path
    sidebar,
    content
])

# callbacks to update page content based on URL and dropdown value
@app.callback(
    Output("page-content", "children"),
    [Input("url", "pathname"), Input("year-dropdown", "value")]
)

def display_page(pathname, selected_year):
    if pathname == "/generation":
        filtered_data = dailydata[dailydata['year'] == selected_year]
        netGen = px.line( # display net generation line graph
            filtered_data, 
            x="Local date", 
            y="D", 
            title=f"Net Generation by Date for {selected_year} (MWh)", 
            labels={"Local date": "Date", "D": "Net Generation (MWh)"},
        )
        netGen.update_layout(
        plot_bgcolor='white')
        netGen.update_xaxes(
            mirror=True,
            ticks='outside',
            showline=True,
            linecolor='black',
            gridcolor='lightgrey'
        )
        netGen.update_yaxes(
            mirror=True,
            ticks='outside',
            showline=True,
            linecolor='black',
            gridcolor='lightgrey'
        )
        netGen.update_traces(line_color='royalblue')
        return html.Div([
            html.Div([
                html.H2(f"Welcome to the Generation and Emissions Analytics Dashboard!"),
                html.P("This dashboard analyzes U.S. Department of Energy power generation and emissions trends over time. Navigate to the sidebar to start exploring the data!"),
            ], style={'background-color': 'whitesmoke', 'padding': '10px', 'border-radius': '5px', 'margin-bottom': '20px'}),
            dcc.Graph(id='generationline', figure=netGen),
            dcc.Graph(id='heatmap', figure=heatmap)
        ], style={"margin-bottom": "2rem"})
    elif pathname == "/emissions":
        emissions_data = dailydata[dailydata['year'] == selected_year]
        emissionsLine = px.line(
            emissions_data,
            x="Local date", 
            y="CO2 Emissions Generated", 
            title=f"Carbon Emissions Generated by Date for {selected_year} (Metric Tons)", 
            labels={"Local date": "Date", "CO2 Emissions Generated": "CO2 Emissions Generated (Metric Tons)"},
        )
        emissionsLine.update_layout(
        plot_bgcolor='white')
        emissionsLine.update_xaxes(
            mirror=True,
            ticks='outside',
            showline=True,
            linecolor='black',
            gridcolor='lightgrey'
        )
        emissionsLine.update_yaxes(
            mirror=True,
            ticks='outside',
            showline=True,
            linecolor='black',
            gridcolor='lightgrey'
        )
        emissionsLine.update_traces(line_color='royalblue') 

        # define line graph with multiple lines for emissions comparison
        comparisonLine = px.line(emissions_data, x='Local date', y=emissions_data.columns[28:31],
                      title=f"Carbon Emissions Generated by Type for {selected_year} (Metric Tons)", 
                      labels={"Local date": "Date"}).update_layout(yaxis_title="CO2 Emissions (Metric Tons)")
        comparisonLine.update_layout(legend_title_text='Type')

        comparisonLine.update_layout(plot_bgcolor='white')

        comparisonLine.update_xaxes(
            mirror=True,
            ticks='outside',
            showline=True,
            linecolor='black',
            gridcolor='lightgrey'
        )

        comparisonLine.update_yaxes(
            mirror=True,
            ticks='outside',
            showline=True,
            linecolor='black',
            gridcolor='lightgrey'
        )
        return html.Div([
            html.Div([
                html.H2(f"Welcome to the Generation and Emissions Analytics Dashboard!"),
                html.P("This dashboard analyzes U.S. Department of Energy power generation and emissions trends over time. Navigate to the sidebar to start exploring the data!"),
            ], style={'background-color': 'whitesmoke', 'padding': '10px', 'border-radius': '5px', 'margin-bottom': '20px'}),
            dcc.Graph(id='generationline', figure=emissionsLine), # plot emissions line
            dcc.Graph(id='line', figure=comparisonLine) # plot comparison line
        ], style={"margin-bottom": "2rem"})


# callback to update heatmap based on selected year
@app.callback(
    Output('heatmap', 'figure'),
    [Input('year-dropdown', 'value')]
)
def update_heatmap(selected_year):
    # filter the data based on selected year
    filtered_hourlydata = hourlydata[hourlydata['date'].dt.year == selected_year]
    filtered_heatmapdata = filtered_hourlydata.pivot_table(index='Hour', columns="month", values='NG', aggfunc='mean')
    filtered_heatmapdata = filtered_heatmapdata.iloc[:-1, :] # remove null values
    filtered_heatmapdata.index = filtered_heatmapdata.index.astype(int)

    # heatmap figure
    updated_heatmap = go.Figure(data=go.Heatmap(z=filtered_heatmapdata.values.tolist(), x=list(filtered_heatmapdata.columns), y=list(filtered_heatmapdata.index), colorscale='Viridis'))
    updated_heatmap.update_layout(title=f"Average Hourly Net Generation Heatmap for {selected_year}", xaxis_title="Month", yaxis_title="Hour",
                          xaxis=dict(tickmode='array', tickvals=list(filtered_heatmapdata.columns), ticktext=['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']),
                          yaxis=dict(tickmode='array', tickvals=list(filtered_heatmapdata.index), ticktext=list(filtered_heatmapdata.index)))
    return updated_heatmap

# run app
if __name__ == "__main__":
    app.run_server(debug=True)
