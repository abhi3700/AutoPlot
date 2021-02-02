import dash_bootstrap_components as dbc

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
            className="area-equipments-layout",
            style={"color": "#ffffff", "font-weight": "bold"}
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
    className="area-equipments-layout",
)