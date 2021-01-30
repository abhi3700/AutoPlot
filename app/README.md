# Coding guides
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