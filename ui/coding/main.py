"""
    Steps:
    ======
    1. card for 'Software name'- "AutoPlot"
    2. Navigation bar for 'Sections' - "DRY ETCH", "WET ETCH",...
    3. Navigation bar for 'Equipments' - ASFE1, ASBE1, REML1, REOX1, REPL1, RESP1
    4. nested Navigation bar for 'Chambers' - 'Ch A', 'Ch C'
    5. nested Navigation bar for 'Charts type' - 'CP', 'ER', 'Unif'
    6. callback functions for pressing dropdown menus - ASFE1-CP Chart, likewise....
"""
import dash
import dash_bootstrap_components as dbc
import dash_html_components as html

app = dash.Dash(external_stylesheets= [dbc.themes.BOOTSTRAP])

# =======================================================================================================
# "AutoPlot" title in card
autoplot_title = dbc.Card(
        dbc.Row(
            [
                html.H3(html.B("AutoPlot"), className="mt-1")
            ],
            justify="center",
            # align="center",
        ),
        color="success",
        inverse=True,
        className="shadow-sm",
        style={
        #     'background-color': "#4caf50"
            "width": "9rem",
            "height": "3rem",
        },
    )

# -------------------------------------------------------------------------------------------------------
# "AutoPlot" subtitle in paragraph
autoplot_subtitle = html.P(html.B("FAB QC Monitor"), className="ml-3",  style={"color": "#9e9e9e", "font-size": "15px"})

# -------------------------------------------------------------------------------------------------------
# "AutoPlot" layout as a col of 2 rows
autoplot_layout = dbc.Col([ 
    dbc.Row(autoplot_title), 
    dbc.Row(autoplot_subtitle)
    ],
    align="center",
    className="m-2"
    )

# =======================================================================================================

app.layout = html.Div(
    [
        autoplot_layout
    ]
)

if __name__ == '__main__':
    app.run_server(debug=True)