"""
    Steps:
    ======
    1. [x] card for 'Software name'- "AutoPlot"
    2. [x] Dropdown menu for 'Sections' - "DRY ETCH", "WET ETCH",...
    3. [x] Navigation bar for 'Equipments' - ASFE1, ASBE1, REML1, REOX1, REPL1, RESP1
    4. [x] nested Navigation bar for 'Chambers' - 'Ch A', 'Ch C'
    5. [x] nested Navigation bar for 'Charts type' - 'CP', 'ER', 'Unif'
    6. [ ] callback functions for pressing dropdown menus - ASFE1-CP Chart, likewise....
"""
import dash
import dash_bootstrap_components as dbc
import dash_html_components as html
from fab_areas.dryetch.layout import area_equipments_layout_dryetch

# external JavaScript files
# external_scripts = [
#     # {'src': 'bootstrap.min.js'},
#     {'src': 'https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js'},
#     # {'src': 'jquery.min.js'},
#     {'src': 'https://code.jquery.com/jquery-3.5.1.min.js'}
# ]

app = dash.Dash(
    # external_scripts= external_scripts,
    external_stylesheets= [
            # {
            #     'href': 'bootstrap.min.css',
            #     'rel': 'stylesheet',
            # },
            # {
            #     'href': 'https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css',
            #     'rel': 'stylesheet',
            # },
            dbc.themes.BOOTSTRAP,
        ]
)

# =======================================================================================================
# "AutoPlot" title in card
autoplot_title = dbc.Card(
        dbc.Row(
            [
                dbc.DropdownMenu(
                    label="AutoPlot",
                    children= [
                        dbc.DropdownMenuItem("Home"),
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
# autoplot_subtitle = html.P(html.B("FAB QC Monitor"), style={"color": "#9e9e9e", "font-size": "15px", "background-color": "#000000"})
# -------------------------------------------------------------------------------------------------------
# add a badge of area
fab_area = dbc.Badge("Dry Etch", id="fab-area", className="ml-1 autoplot-shadow", pill=True, style={"color": "#424242", "background-color": "#ff8f00", "font-size": "14px"})
# -------------------------------------------------------------------------------------------------------
# "AutoPlot" layout as a col of 2 rows
autoplot_layout = dbc.Col(
    [ 
        dbc.Row(autoplot_title), 
        dbc.Row(
            [
                html.Span(
                    html.H1(
                    [
                        html.B("FAB QC Monitor  | "),            # autoplot_subtitle
                        fab_area                                # fab_area
                    ],                               
                    className="ml-2 mt-1 autoplot-subtitle"
                    )
                )
            ],
        )
    ],
    align="center",
    className="m-2"
    )

# callback for changing fab_area text (badge) by clicking dropdown menu

# =======================================================================================================
# "buttongroup with nested dropdownmenu" for area equipments

area_equipments_layout = area_equipments_layout_dryetch
# =======================================================================================================
app.layout = html.Div(
    [
        autoplot_layout,
        area_equipments_layout,
    ],
)

if __name__ == '__main__':
    app.run_server(debug=True)