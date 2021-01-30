"""
    Steps:
    ======
    1. [x] card for 'Software name'- "AutoPlot"
    2. [ ] Dropdown menu for 'Sections' - "DRY ETCH", "WET ETCH",...
    3. [ ] Navigation bar for 'Equipments' - ASFE1, ASBE1, REML1, REOX1, REPL1, RESP1
    4. [ ] nested Navigation bar for 'Chambers' - 'Ch A', 'Ch C'
    5. [ ] nested Navigation bar for 'Charts type' - 'CP', 'ER', 'Unif'
    6. [ ] callback functions for pressing dropdown menus - ASFE1-CP Chart, likewise....
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
                dbc.DropdownMenu(
                    label="AutoPlot",
                    children= [
                        dbc.DropdownMenuItem("CMP"),
                        dbc.DropdownMenuItem("Diffusion"),
                        dbc.DropdownMenuItem("Dry Etch"),
                        dbc.DropdownMenuItem("Implant"),
                        dbc.DropdownMenuItem("Photo"),
                        dbc.DropdownMenuItem("Thin Film"),
                        dbc.DropdownMenuItem("Wet Etch"),
                        dbc.DropdownMenuItem("Yield"),
                    ],
                    bs_size="lg",
                    color="success",
                )
            ],
            justify="center",
        ),
        color="success",
        inverse=True,
        className="autoplot-shadow",
        style={
        #     'background-color': "#4caf50"
            "width": "8rem",
            "height": "3rem",
        },
    )


# -------------------------------------------------------------------------------------------------------
# "AutoPlot" subtitle in paragraph
autoplot_subtitle = html.P(html.B("FAB QC Monitor"), className="ml-2",  style={"color": "#9e9e9e", "font-size": "15px"})

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