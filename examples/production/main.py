import plotly as py
import plotly.graph_objs as go

import pandas as pd

# -----------------------------------------------------------------------
# input vars for MTRS1
excel_file_directory = 'Average Products Lot yield Aeff only 2020.xlsx'

trace_all_marker_border_color = '#ffffff'

x_label = 'Effective Area (in mm2)'
y_label = 'Lot IDs'
z_label = 'Lot Yield (in %)'
plot_title = "Average Products Lot yield Aeff only 2020"
plot_filename = 'Average_Products_Lot_yield_Aeff_only_2020.html'


# def scatter_plot_3d(x, y, z, remarks_1, remarks_2):
# def scatter_plot_3d(
#                     x1, y1, z1, remarks_1,
#                     ):
#     trace1 = go.Scatter3d(
#                 x= df1[cols[0]].tolist(),
#                 y= df1[cols[3]].tolist(),
#                 z= df1[cols[4]].tolist(),
#                 name=df1[cols[1]].tolist()[0],
#                 mode='markers',
#                 line = dict(
#                         color = trace1_line_color,
#                         width = 2),
#                 marker = dict(
#                         color = trace1_marker_color,
#                         size = 8,
#                         line = dict(
#                             color = trace1_marker_border_color,
#                             width = 0.5),
#                         ),
#                 text = df1[cols[2]].tolist()
#                 )
    

#     layout = dict(
#                 title = plot_title,
#                 xaxis = dict(title= x_label),
#                 yaxis = dict(title= y_label)
#                 )

#     data = [trace1]
#     fig = dict(data= data, layout= layout)
#     py.offline.plot(fig, filename= plot_filename)



# ----------------------------------------------------------------------
# Read data from a csv
df = pd.read_excel(excel_file_directory, sheet_name= "Average Products yield Aeff onl", skiprows= 0)
# print(df)
cols = ["Aeff (mm2)", "Device_ID", "Device_Name", "LotID", "Lot Yield (%)"]

area_eff = df[cols[0]].tolist()
device_ids = df[cols[1]].tolist()
device_names = df[cols[2]].tolist()
lot_ids = df[cols[3]].tolist()
lot_yield_percent = df[cols[4]].tolist()

df1 = df[df[cols[1]] == 'SC1001-0']
df2 = df[df[cols[1]] == 'SC1002-1']
df3 = df[df[cols[1]] == 'SC1004-2']
df4 = df[df[cols[1]] == 'SC1005_1']
df5 = df[df[cols[1]] == 'SC1009-0']
df6 = df[df[cols[1]] == 'SC1012-0']
df7 = df[df[cols[1]] == 'SC1013-0']
df8 = df[df[cols[1]] == 'SC1021-0']
df9 = df[df[cols[1]] == 'SC1022-0']
df10 = df[df[cols[1]] == 'SC1023-0']
df11 = df[df[cols[1]] == 'SC1106']
df12 = df[df[cols[1]] == 'SC1117-0']
df13 = df[df[cols[1]] == 'SC1118-0']
df14 = df[df[cols[1]] == 'SC1120']
df15 = df[df[cols[1]] == 'SC1121_0']
df16 = df[df[cols[1]] == 'SC1123']
df17 = df[df[cols[1]] == 'SC1123-0']
df18 = df[df[cols[1]] == 'SC1124_0']
df19 = df[df[cols[1]] == 'SC1127']
df20 = df[df[cols[1]] == 'SC1129-0']
df21 = df[df[cols[1]] == 'SC1130_0']
df22 = df[df[cols[1]] == 'SC1203-0T1']
df23 = df[df[cols[1]] == 'SC1216_0']
df24 = df[df[cols[1]] == 'SC1218']
df25 = df[df[cols[1]] == 'SC1221']
df26 = df[df[cols[1]] == 'SC1229-0']
df27 = df[df[cols[1]] == 'SC1232-0']
df28 = df[df[cols[1]] == 'SC1239-0']
df29 = df[df[cols[1]] == 'SC1240']
df30 = df[df[cols[1]] == 'SC1301-1']
df31 = df[df[cols[1]] == 'SC1402-0']
df32 = df[df[cols[1]] == 'SC1408-0']
df33 = df[df[cols[1]] == 'SC1506_0']
df34 = df[df[cols[1]] == 'SC1507_0']
df35 = df[df[cols[1]] == 'SC1702-0']
df36 = df[df[cols[1]] == 'SC1703_0']
df37 = df[df[cols[1]] == 'SC9005-0']
df38 = df[df[cols[1]] == 'SC9007-0']
df39 = df[df[cols[1]] == 'SC9015_0']
df40 = df[df[cols[1]] == 'SCL.TTV']

# create additional column for each dataframe
df1['marker_color'] = ['#264653' for i in range(len(df1[cols[0]].tolist()))]
df2['marker_color'] = ['#2a9d8f' for i in range(len(df2[cols[0]].tolist()))]
df3['marker_color'] = ['#e9c46a' for i in range(len(df3[cols[0]].tolist()))]
df4['marker_color'] = ['#f4a261' for i in range(len(df4[cols[0]].tolist()))]
df5['marker_color'] = ['#e76f51' for i in range(len(df5[cols[0]].tolist()))]
df6['marker_color'] = ['#e63946' for i in range(len(df6[cols[0]].tolist()))]
df7['marker_color'] = ['#a8dadc' for i in range(len(df7[cols[0]].tolist()))]
df8['marker_color'] = ['#a8dadc' for i in range(len(df8[cols[0]].tolist()))]
df9['marker_color'] = ['#457b9d' for i in range(len(df9[cols[0]].tolist()))]
df10['marker_color'] = ['#1d3557' for i in range(len(df10[cols[0]].tolist()))]
df11['marker_color'] = ['#ffb4a2' for i in range(len(df11[cols[0]].tolist()))]
df12['marker_color'] = ['#e5989b' for i in range(len(df12[cols[0]].tolist()))]
df13['marker_color'] = ['#b5838d' for i in range(len(df13[cols[0]].tolist()))]
df14['marker_color'] = ['#6d6875' for i in range(len(df14[cols[0]].tolist()))]
df15['marker_color'] = ['#cb997e' for i in range(len(df15[cols[0]].tolist()))]
df16['marker_color'] = ['#eddcd2' for i in range(len(df16[cols[0]].tolist()))]
df17['marker_color'] = ['#ddbea9' for i in range(len(df17[cols[0]].tolist()))]
df18['marker_color'] = ['#a5a58d' for i in range(len(df18[cols[0]].tolist()))]
df19['marker_color'] = ['#b7b7a4' for i in range(len(df19[cols[0]].tolist()))]
df20['marker_color'] = ['#023e8a' for i in range(len(df20[cols[0]].tolist()))]
df21['marker_color'] = ['#0077b6' for i in range(len(df21[cols[0]].tolist()))]
df22['marker_color'] = ['#00b4d8' for i in range(len(df22[cols[0]].tolist()))]
df23['marker_color'] = ['#06d6a0' for i in range(len(df23[cols[0]].tolist()))]
df24['marker_color'] = ['#118ab2' for i in range(len(df24[cols[0]].tolist()))]
df25['marker_color'] = ['#606c38' for i in range(len(df25[cols[0]].tolist()))]
df26['marker_color'] = ['#ff006e' for i in range(len(df26[cols[0]].tolist()))]
df27['marker_color'] = ['#ffb703' for i in range(len(df27[cols[0]].tolist()))]
df28['marker_color'] = ['#80b918' for i in range(len(df28[cols[0]].tolist()))]
df29['marker_color'] = ['#2ec4b6' for i in range(len(df29[cols[0]].tolist()))]
df30['marker_color'] = ['#fee440' for i in range(len(df30[cols[0]].tolist()))]
df31['marker_color'] = ['#9b5de5' for i in range(len(df31[cols[0]].tolist()))]
df32['marker_color'] = ['#403d39' for i in range(len(df32[cols[0]].tolist()))]
df33['marker_color'] = ['#511f73' for i in range(len(df33[cols[0]].tolist()))]
df34['marker_color'] = ['#143601' for i in range(len(df34[cols[0]].tolist()))]
df35['marker_color'] = ['#3d348b' for i in range(len(df35[cols[0]].tolist()))]
df36['marker_color'] = ['#390099' for i in range(len(df36[cols[0]].tolist()))]
df37['marker_color'] = ['#ff1654' for i in range(len(df37[cols[0]].tolist()))]
df38['marker_color'] = ['#b1cf5f' for i in range(len(df38[cols[0]].tolist()))]
df39['marker_color'] = ['#393d3f' for i in range(len(df39[cols[0]].tolist()))]
df40['marker_color'] = ['#ffbf81' for i in range(len(df40[cols[0]].tolist()))]

# print(df40)

dfs = [
        df1, df2, df3, df4, df5, df6, df7, df8, df9, df10,
        df11, df12, df13, df14, df15, df16, df17, df18, df19, df20,
        df21, df22, df23, df24, df25, df26, df27, df28, df29, df30,
        df31, df32, df33, df34, df35, df36, df37, df38, df39, df40
      ]

# print(area_eff)
# print(lot_ids)
# print(lot_yield_percent)
# print(df1[cols[1]].tolist())
# print(df1['marker_color'].tolist()[0])


data = []
for df in  dfs:
    c = df['marker_color'].tolist()[0]
    trace = go.Scatter3d(
                    x= df[cols[0]].tolist(),
                    y= df[cols[3]].tolist(),
                    z= df[cols[4]].tolist(),
                    name=df[cols[1]].tolist()[0],
                    mode='markers',
                    marker = dict(
                            color = c,
                            size = 8,
                            line = dict(
                                color = trace_all_marker_border_color,
                                width = 0.5),
                            ),
                    text = df[cols[2]].tolist()
                )
    data.append(trace)


layout = dict(
            title = plot_title,
            xaxis = dict(title= x_label),
            yaxis = dict(title= y_label)
            )


fig = dict(data= data, layout= layout)
py.offline.plot(fig, filename= plot_filename)
