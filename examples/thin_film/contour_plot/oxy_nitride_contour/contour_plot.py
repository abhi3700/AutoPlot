import plotly as py
import plotly.graph_objs as go
import pandas as pd
from input import *


# ===========================================================================================================
# Read data from a Excel
df = pd.read_excel(excel_file_directory, sheet_name= sht_name)

# define x, y, z
x = df['x'].tolist()
y = df['y'].tolist()
z = df['thickness'].tolist()

# ===========================================================================================================
"""
"Description": Draw contour plot
"x": x value
"y": y value
"z": z value
"fname": filename for the contour plot
"tname": title for contour plot
"num_data_point": number of data points
"""
def contour_plot(x, y, z, fname, tname, num_data_point):
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
                title = tname,
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
                    text= list(range(1, num_data_point + 1))
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
    py.offline.plot(fig, filename= fname + '.html')


# ===========================================================================================================
# Draw contour plot
contour_plot(
    x = x,
    y = y,
    z = z,
    fname = fname,
    tname = tname, 
    num_data_point = num_data_point
    )

