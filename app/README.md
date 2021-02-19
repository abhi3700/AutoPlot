# App

## Features
* [x] card for 'Software name'- "AutoPlot"
* [x] Dropdown menu for 'Sections' - "DRY ETCH", "WET ETCH",...
* [x] Navigation bar for 'Equipments' - ASFE1, ASBE1, REML1, REOX1, REPL1, RESP1
* [x] nested Navigation bar for 'Chambers' - 'Ch A', 'Ch C'
* [x] nested Navigation bar for 'Charts type' - 'CP', 'ER', 'Unif'
* [ ] callback functions for pressing multi-level dropdown menus items-  ASFE1-CP Chart, etc...
* [ ] Dashboard
    - [ ] no. of charts representation in pie chart for each equipments
        - total 59 charts
    - [ ] QC frequency for all equipments
    - [ ] QC procedure for each equipment

## Coding guides
* buttongroup with dropdown
```py
"""
    Buttongroup with dropdown menu
"""
import dash
import dash_bootstrap_components as dbc
import dash_html_components as html

app = dash.Dash(external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = html.Div(
    [
        dbc.ButtonGroup(
            [
            dbc.Button("First"), 
            dbc.Button("Second"), 
            dbc.DropdownMenu([dbc.DropdownMenuItem("Item 1"),dbc.DropdownMenuItem("Item 2")],
                label="Dropdown",
                group=True)
            ],
            ),
    ]
)

if __name__ == '__main__':
    app.run_server(debug=True)
```

* direction dropdown
```py
"""
    dropdown menu direction
"""
import dash
import dash_bootstrap_components as dbc
import dash_html_components as html

app = dash.Dash(external_stylesheets=[dbc.themes.BOOTSTRAP])

items = [
    dbc.DropdownMenuItem("First"),
    dbc.DropdownMenuItem(divider=True),
    dbc.DropdownMenuItem("Second"),
]

dropdown = dbc.Row(
    [
        dbc.Col(dbc.DropdownMenu( items, label="Dropdown (default)", direction= "down", className="mb-3" ), width="auto"),
        dbc.Col(dbc.DropdownMenu( items, label="Dropleft", direction= "left", className="mb-3" ), width="auto"),
        dbc.Col(dbc.DropdownMenu( items, label="DropRight", direction= "right", className="mb-3" ), width="auto"),
        dbc.Col(dbc.DropdownMenu( items, label="Dropup", direction= "up", className="mb-3" ), width="auto"),
    ],
    justify="between",	# The justify-content property aligns the flexible container's items when the items do not use all available space on the main-axis (horizontally). params like flex-start|flex-end|center|space-between|space-around|space-evenly|initial|inherit;
    # reference: https://www.w3schools.com/cssref/css3_pr_justify-content.asp
)

app.layout = html.Div([dropdown])


if __name__ == '__main__':
    app.run_server(debug=True)
```

