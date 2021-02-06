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
# dash libs
import dash
import dash_bootstrap_components as dbc
import dash_html_components as html
import dash_core_components as dcc
from dash.dependencies import Input, Output, State

# area layouts
from fab_areas.home.layout import area_equipments_layout_home
from fab_areas.cmp.layout import area_equipments_layout_cmp
from fab_areas.diffusion.layout import area_equipments_layout_diffusion
from fab_areas.dryetch.layout import area_equipments_layout_dryetch
from fab_areas.implant.layout import area_equipments_layout_implant
from fab_areas.photo.layout import area_equipments_layout_photo
from fab_areas.thinfilm.layout import area_equipments_layout_thinfilm
from fab_areas.wetetch.layout import area_equipments_layout_wetetch
from fab_areas.yieldtdd.layout import area_equipments_layout_yield

# area equipments
# Dry Etch
from fab_areas.dryetch.equipments.ASFE1.ASFE1 import asfe1_cp_chart, asfe1_er_chart, asfe1_unif_chart
from fab_areas.dryetch.equipments.ASBE1.ASBE1 import asbe1_cp_chart, asbe1_er_chart, asbe1_unif_chart
from fab_areas.dryetch.equipments.REML1.REML1 import reml1a_cp_chart, reml1c_cp_chart, reml1a_pr_er_chart, reml1a_pr_unif_chart, reml1c_pr_er_chart, reml1c_pr_unif_chart

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

# =======================================================================================================
# "buttongroup with nested dropdownmenu" for area equipments

area_equipments_layout = html.Div(
        area_equipments_layout_dryetch, 
        id= "area-equipment-layout",
    )


# callback for changing "fab_area text (badge)" & "area_equipment_layout" by clicking dropdown menu item
@app.callback(
    [
        Output('fab-area', 'children'),
        Output('area-equipment-layout', 'children'),
    ],
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
    area_layout = { 
                    'home-id': area_equipments_layout_home,
                    'cmp-id': area_equipments_layout_cmp,
                    'diffusion-id': area_equipments_layout_diffusion,
                    'dryetch-id': area_equipments_layout_dryetch,
                    'implant-id': area_equipments_layout_implant,
                    'thinfilm-id': area_equipments_layout_thinfilm,
                    'photo-id': area_equipments_layout_photo,
                    'wetetch-id': area_equipments_layout_wetetch,
                    'yield-id': area_equipments_layout_yield,
                }


    ctx = dash.callback_context
    if not ctx.triggered:
        # out_text = "Home"
        out_text = "Dry Etch"
        out_layout = area_equipments_layout_dryetch
    else:
        button_id = ctx.triggered[0]['prop_id'].split(".")[0]
        # print(ctx.triggered[0])
        # print(button_id)
        out_text = fabarea_label[button_id]
        out_layout = area_layout[button_id]

    return out_text, out_layout

# =======================================================================================================
# container for graph
chart = html.Div( 
            [
                dcc.Graph(id='area-equip-ch-chart', className= "full-height")
            ],
            style={
                "background-color": "#EF9A9A",
                "height": "auto",
            },
            className="full-height",
        )


@app.callback(
    Output('area-equip-ch-chart', 'figure'),
    [
        # ASFE1
        Input('asfe1-cp-chart', 'n_clicks'),
        Input('asfe1-er-chart', 'n_clicks'),
        Input('asfe1-unif-chart', 'n_clicks'),
        # ASBE1
        Input('asbe1-cp-chart', 'n_clicks'),
        Input('asbe1-er-chart', 'n_clicks'),
        Input('asbe1-unif-chart', 'n_clicks'),
        # REML1
        Input('reml1a-cp-chart', 'n_clicks'),
        Input('reml1a-pr-er-chart', 'n_clicks'),
        Input('reml1a-pr-unif-chart', 'n_clicks'),
        # Input('reml1a-al-er-chart', 'n_clicks'),
        # Input('reml1a-al-unif-chart', 'n_clicks'),
        Input('reml1c-cp-chart', 'n_clicks'),
        Input('reml1c-pr-er-chart', 'n_clicks'),
        Input('reml1c-pr-unif-chart', 'n_clicks'),
    ]
)
def update_chart(*args):
    chart_func = {
            # ASFE1
            'asfe1-cp-chart': asfe1_cp_chart(),
            'asfe1-er-chart': asfe1_er_chart(),
            'asfe1-unif-chart': asfe1_unif_chart(),
            # ASBE1
            'asbe1-cp-chart': asbe1_cp_chart(),
            'asbe1-er-chart': asbe1_er_chart(),
            'asbe1-unif-chart': asbe1_unif_chart(),
            # REML1
            'reml1a-cp-chart': reml1a_cp_chart(),
            'reml1a-pr-er-chart': reml1a_pr_er_chart(),
            'reml1a-pr-unif-chart': reml1a_pr_unif_chart(),
            # 'reml1a-al-er-chart': reml1a_al_er_chart(),
            # 'reml1a-al-unif-chart': reml1a_al_unif_chart(),
            'reml1c-cp-chart': reml1c_cp_chart(),
            'reml1c-pr-er-chart': reml1c_pr_er_chart(),
            'reml1c-pr-unif-chart': reml1c_pr_unif_chart(),
    }
    ctx = dash.callback_context

    if not ctx.triggered:
        fig = {}
    else:
        button_id = ctx.triggered[0]['prop_id'].split(".")[0]
        fig = chart_func[button_id]
        
    return fig
# =======================================================================================================
app.layout = html.Div(
    [
        autoplot_layout,
        area_equipments_layout,
        chart,
    ],
)

if __name__ == '__main__':
    app.run_server(debug=True)