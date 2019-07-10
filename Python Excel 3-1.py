from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series, LineChart, ScatterChart
from openpyxl.styles import Font, Color, colors

wb = Workbook()
ws = wb.active
for i in range(10):
    ws.append([i])

# drawing a graph
values = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=10)
# chart = LineChart()
chart = BarChart()
ws.add_chart(chart, "A15")
chart.title = "Chart"
chart.y_axis.title = 'Size'
chart.x_axis.title = 'Test Number'
chart.add_data(values)

s1 = chart.series[0]
s1.marker.symbol = "triangle"



wb.save("Chart-1.xlsx")
