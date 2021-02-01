"""
    Steps:
    ======
    1. [x] card for 'Software name'- "AutoPlot"
    2. [x] Dropdown menu for 'Sections' - "DRY ETCH", "WET ETCH",...
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
autoplot_subtitle = html.P(html.B("FAB QC Monitor"), style={"color": "#9e9e9e", "font-size": "15px", "background-color": "#000000"})
# -------------------------------------------------------------------------------------------------------
# add a badge of area
fab_area = dbc.Badge("Dry Etch", id="fab-area", className="ml-1", pill=True, style={"color": "#ffffff", "background-color": "#ff8f00", "font-size": "14px"})
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
                        html.B("FAB QC Monitor | "),            # autoplot_subtitle
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
area_equipments_layout_dryetch = dbc.ButtonGroup(
    [
        dbc.DropdownMenu(
            [
                dbc.DropdownMenuItem("CP Chart", id="asfe1-cp-chart"),
                dbc.DropdownMenuItem("ER Chart", id="asfe1-er-chart"),
                dbc.DropdownMenuItem("Unif Chart", id="asfe1-unif-chart"),
            ],
            label="ASFE1",
            group=True,
            color="#00e676",
            className="area-equipments-layout text-white",
            style={"color": "#ffffff"}
        ),
        dbc.DropdownMenu(
            [
                dbc.DropdownMenuItem("CP Chart", id="asbe1-cp-chart"),
                dbc.DropdownMenuItem("ER Chart", id="asbe1-er-chart"),
                dbc.DropdownMenuItem("Unif Chart", id="asbe1-unif-chart"),
            ],
            label="ASBE1",
            group=True,
            color="#00e676",
            className="area-equipments-layout",
        ),
        dbc.DropdownMenu(
            [
                dbc.ButtonGroup(
                    [
                        dbc.DropdownMenu(
                            [
                                dbc.DropdownMenuItem("CP Chart", id="reml1a-cp-chart"),
                                dbc.DropdownMenuItem("PR ER Chart", id="reml1a-pr-er-chart"),
                                dbc.DropdownMenuItem("PR Unif Chart", id="reml1a-pr-unif-chart"),
                                dbc.DropdownMenuItem("Al ER Chart", id="reml1a-al-er-chart"),
                                dbc.DropdownMenuItem("Al Unif Chart", id="reml1a-al-unif-chart"),
                            ],
                            label="Ch A",
                            direction="right",
                            group=True,
                            color="#ffffff",
                        ),
                        dbc.DropdownMenu(
                            [
                                dbc.DropdownMenuItem("CP Chart", id="reml1c-cp-chart"),
                                dbc.DropdownMenuItem("PR ER Chart", id="reml1c-pr-er-chart"),
                                dbc.DropdownMenuItem("PR Unif Chart", id="reml1c-pr-unif-chart"),
                            ],
                            label="Ch C",
                            direction="right",
                            group=True,
                            color="#ffffff",
                        ),
                    ],
                    vertical=True,
                    className="area-equipments-layout",
                )
            ],
            label="REML1",
            group=True,
            color="#00e676",
            className="area-equipments-layout",
        ),
        dbc.DropdownMenu(
            [
                dbc.ButtonGroup(
                    [
                        dbc.DropdownMenu(
                            [
                                dbc.DropdownMenuItem("CP Chart", id="reox1a-cp-chart"),
                                dbc.DropdownMenuItem("TEOS ER Chart", id="reox1a-teos-er-chart"),
                                dbc.DropdownMenuItem("TEOS Unif Chart", id="reox1a-teos-unif-chart"),
                                dbc.DropdownMenuItem("SiN ER Chart", id="reox1a-sin-er-chart"),
                                dbc.DropdownMenuItem("SiN Unif Chart", id="reox1a-sin-unif-chart"),
                            ],
                            label="Ch A",
                            direction="right",
                            group=True,
                            color="#ffffff",
                        ),
                        dbc.DropdownMenu(
                            [
                                dbc.DropdownMenuItem("CP Chart", id="reox1b-cp-chart"),
                                dbc.DropdownMenuItem("BPSG_CS ER Chart", id="reox1b-bpsgcs-er-chart"),
                                dbc.DropdownMenuItem("BPSG_CS Unif Chart", id="reox1b-bpsgcs-unif-chart"),
                                dbc.DropdownMenuItem("SiN_CS ER Chart", id="reox1b-sincs-er-chart"),
                                dbc.DropdownMenuItem("SiN_CS Unif Chart", id="reox1b-sincs-unif-chart"),
                                dbc.DropdownMenuItem("TEOS_VIA ER Chart", id="reox1b-teosvia-er-chart"),
                                dbc.DropdownMenuItem("TEOS_VIA Unif Chart", id="reox1b-teosvia-unif-chart"),
                                dbc.DropdownMenuItem("ARC ER Chart", id="reox1b-arc-er-chart"),
                                dbc.DropdownMenuItem("ARC Unif Chart", id="reox1b-arc-unif-chart"),
                            ],
                            label="Ch B",
                            direction="right",
                            group=True,
                            color="#ffffff",
                        ),
                        dbc.DropdownMenu(
                            [
                                dbc.DropdownMenuItem("CP Chart", id="reox1c-cp-chart"),
                                dbc.DropdownMenuItem("ARC ER Chart", id="reox1c-arc-er-chart"),
                                dbc.DropdownMenuItem("ARC Unif Chart", id="reox1c-arc-unif-chart"),
                                dbc.DropdownMenuItem("TEOS ER Chart", id="reox1c-teos-er-chart"),
                                dbc.DropdownMenuItem("TEOS Unif Chart", id="reox1c-teos-unif-chart"),
                            ],
                            label="Ch C",
                            direction="right",
                            group=True,
                            color="#ffffff",
                        ),
                    ],
                    vertical=True,
                    className="area-equipments-layout",
                )
            ],
            label="REOX1",
            group=True,
            color="#00e676",
            className="area-equipments-layout",
        ),
        dbc.DropdownMenu(
            [
                dbc.ButtonGroup(
                    [
                        dbc.DropdownMenu(
                            [
                                dbc.DropdownMenuItem("CP Chart", id="repl1a-cp-chart"),
                                dbc.DropdownMenuItem("Nit ER Chart", id="repl1a-nit-er-chart"),
                                dbc.DropdownMenuItem("Nit Unif Chart", id="repl1a-nit-unif-chart"),
                                dbc.DropdownMenuItem("Poly ER Chart", id="repl1a-poly-er-chart"),
                                dbc.DropdownMenuItem("Poly Unif Chart", id="repl1a-poly-unif-chart"),
                            ],
                            label="Ch A",
                            direction="right",
                            group=True,
                            color="#ffffff",
                        ),
                        dbc.DropdownMenu(
                            [
                                dbc.DropdownMenuItem("CP Chart", id="repl1b-cp-chart"),
                                dbc.DropdownMenuItem("Nit ER Chart", id="repl1b-nit-er-chart"),
                                dbc.DropdownMenuItem("Nit Unif Chart", id="repl1b-nit-unif-chart"),
                                dbc.DropdownMenuItem("Poly ER Chart", id="repl1b-poly-er-chart"),
                                dbc.DropdownMenuItem("Poly Unif Chart", id="repl1b-poly-unif-chart"),
                            ],
                            label="Ch B",
                            direction="right",
                            group=True,
                            color="#ffffff",
                        ),
                    ],
                    vertical=True,
                    className="area-equipments-layout",
                )
            ],
            label="REPL1",
            group=True,
            color="#00e676",
            className="area-equipments-layout",
        ),
        dbc.DropdownMenu(
            [
                dbc.ButtonGroup(
                    [
                        dbc.DropdownMenu(
                            [
                                dbc.DropdownMenuItem("CP Chart", id="resp1a-cp-chart"),
                                dbc.DropdownMenuItem("TEOS-1st ER Chart", id="resp1a-teos1st-er-chart"),
                                dbc.DropdownMenuItem("TEOS-1st Unif Chart", id="resp1a-teos1st-unif-chart"),
                                dbc.DropdownMenuItem("TEOS-2nd ER Chart", id="resp1a-teos2nd-er-chart"),
                                dbc.DropdownMenuItem("TEOS-2nd Unif Chart", id="resp1a-teos2nd-unif-chart"),
                                dbc.DropdownMenuItem("SiN-1st ER Chart", id="resp1a-sin1st-er-chart"),
                                dbc.DropdownMenuItem("SiN-1st Unif Chart", id="resp1a-sin1st-unif-chart"),
                                dbc.DropdownMenuItem("SiN-2nd ER Chart", id="resp1a-sin2nd-er-chart"),
                                dbc.DropdownMenuItem("SiN-2nd Unif Chart", id="resp1a-sin2nd-unif-chart"),
                            ],
                            label="Ch A",
                            direction="right",
                            group=True,
                            color="#ffffff",
                        ),
                        dbc.DropdownMenu(
                            [
                                dbc.DropdownMenuItem("CP Chart", id="resp1b-cp-chart"),
                                dbc.DropdownMenuItem("BARC ER Chart", id="resp1b-barc-er-chart"),
                                dbc.DropdownMenuItem("BARC Unif Chart", id="resp1b-barc-unif-chart"),
                                dbc.DropdownMenuItem("PR ER Chart", id="resp1b-pr-er-chart"),
                                dbc.DropdownMenuItem("PR Unif Chart", id="resp1b-pr-unif-chart"),
                                dbc.DropdownMenuItem("TEOS ER Chart", id="resp1b-teos-er-chart"),
                                dbc.DropdownMenuItem("TEOS Unif Chart", id="resp1b-teos-unif-chart"),
                                dbc.DropdownMenuItem("SiN ER Chart", id="resp1b-sin-er-chart"),
                                dbc.DropdownMenuItem("SiN Unif Chart", id="resp1b-sin-unif-chart"),
                            ],
                            label="Ch B",
                            direction="right",
                            group=True,
                            color="#ffffff",
                        ),
                    ],
                    vertical=True,
                    className="area-equipments-layout",
                )
            ],
            label="RESP1",
            group=True,
            color="#00e676",
            className="area-equipments-layout",
        ),
    ],
    className="area-equipments-layout text-white",
)

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