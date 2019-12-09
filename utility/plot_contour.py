"""
"Description": This function plots Contour (2-D) chart of ER for different layers.
"x": x-coordinate for ER Chart
"y": y-coordinate for ER Chart
"z": z values for ER Chart
"fname": filename for ER Chart
"tname": title name for ER Chart
"date": date for ER Chart
"""
def contour_plot(x, y, z, fname, tname, date):
    trace1 = go.Contour(
                x= x,
                y= y,
                z= z,
                contours=dict(
                            # coloring ='heatmap',
                            showlabels = True, # show labels on contours
                            labelfont = dict( # label font properties
                                size = 12,
                                color = 'white',
                               )
                        ),
                )
    

    layout = dict(
                title = 'Contour plot of ' + tname + ' on ' + date + ' for REPL1 Ch A',
                xaxis = dict(
                            title= 'x-axis (in mm)',
                            # type='linear',
                            range= [-100, 100]
                            ),
                yaxis = dict(
                            title= 'y-axis (in mm)',
                            # type='linear',
                            range= [-100, 100]
                            ),
                autosize= True,
                width=800, height=800,

                )

    data = [trace1]
    # fig = dict(data= data, layout= layout)
    fig = go.Figure(data= data, layout= layout)     # an `graph_objs` object created
    fig.add_trace(      # added a scatter plot for showing the markers onto the contour plot
            go.Scatter(
                    mode='markers',
                    x= x,
                    y= y,
                    opacity=0.8,
                    marker=dict(
                        color='#ffffff',
                        size=10,
                        line=dict(
                            color='#ffffff',
                            width= 0.5
                        ),
                        symbol = 'x'
                    ),
                    showlegend=False,
                    text= list(range(1,14))
            )        
        )
    fig.update_layout(      # added a unfilled circle from (-100, -100) to (100, 100) diagonally
        shapes=[
            # unfilled circle
            go.layout.Shape(
                type="circle",
                xref="x",
                yref="y",
                x0= -100,
                y0= -100,
                x1= 100,
                y1= 100,
                line_color="#212121",
            ),
        ]
    )
    py.offline.plot(fig, filename= fname)
