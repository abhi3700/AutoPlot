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
import dash_core_components as dcc
from dash.dependencies import Input, Output, State
from fab_areas.dryetch.layout import area_equipments_layout_dryetch
from fab_areas.dryetch.equipments.ASFE1.ASFE1 import cp_chart, er_chart, unif_chart

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
                        dbc.DropdownMenuItem("Home", id= "home-id"),
                        dbc.DropdownMenuItem("CMP", id= "cmp-id"),
                        dbc.DropdownMenuItem("Diffusion", id= "diffusion-id"),
                        dbc.DropdownMenuItem("Dry Etch", id= "dryetch-id"),
                        dbc.DropdownMenuItem("Implant", id= "implant-id"),
                        dbc.DropdownMenuItem("Photo", id= "photo-id"),
                        dbc.DropdownMenuItem("Thin Film", id= "thinfilm-id"),
                        dbc.DropdownMenuItem("Wet Etch", id= "wetetch-id"),
                        dbc.DropdownMenuItem("Yield", id= "yield-id"),
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
fab_area = dbc.Badge(children="Home", id="fab-area", className="ml-1 autoplot-shadow", pill=True, style={"color": "#424242", "background-color": "#ff8f00", "font-size": "14px"})
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
    className="m-2",
    )

# callback for changing fab_area text (badge) by clicking dropdown menu item
@app.callback(
    Output('fab-area', 'children'),
    [
        Input('home-id', 'n_clicks'),
        Input('cmp-id', 'n_clicks'),
        Input('diffusion-id', 'n_clicks'),
        Input('dryetch-id', 'n_clicks'),
        Input('implant-id', 'n_clicks'),
        Input('photo-id', 'n_clicks'),
        Input('thinfilm-id', 'n_clicks'),
        Input('wetetch-id', 'n_clicks'),
        Input('yield-id', 'n_clicks'),
    ]
)
def update_fabarea_badge(*args):
    # create a dict for mapping the button-id with desired label
    fabarea_label = { 
                        'home-id': 'Home',
                        'cmp-id': 'CMP',
                        'diffusion-id': 'Diffusion',
                        'dryetch-id': 'Dry Etch',
                        'implant-id': 'Implant',
                        'thinfilm-id': 'Thin Film',
                        'photo-id': 'Photo',
                        'wetetch-id': 'Wet Etch',
                        'yield-id': 'Yield',
                    }

    ctx = dash.callback_context
    if not ctx.triggered:
        out_text = "Home"
    else:
        button_id = ctx.triggered[0]['prop_id'].split(".")[0]
        # print(ctx.triggered[0])
        # print(button_id)
        out_text = fabarea_label[button_id]

    return out_text
# =======================================================================================================
# "buttongroup with nested dropdownmenu" for area equipments

area_equipments_layout = area_equipments_layout_dryetch


# callback for changing area_equipments_layout (button_group) by change of badge text
# =======================================================================================================
# container for graph
"""
generate_chart = dcc.Graph(id='area-equip-ch-chart', figure={})

@app.callback(
    Output('area-equip-ch-chart', 'figure'),
    Input('asfe1-cp-chart', 'toggle')
)
def creatify_chart(input_toggle):
    if input_toggle:
        fig = cp_chart()
        return fig
    else:
        return None

"""
# =======================================================================================================
app.layout = html.Div(
    [
        autoplot_layout,
        area_equipments_layout,
        # html.Div(generate_chart),
    ],
)

if __name__ == '__main__':
    app.run_server(debug=True)