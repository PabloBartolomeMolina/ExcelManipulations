from openpyxl.chart import Reference, LineChart


# Create a line chart, using as an input the worksheet containing the data that will include the chart.
# To take the good parameters for the chart, we need to take as input the definition of the ranges of cells.
def line_chart_create(wb, ws, filename, content, range_axis):
    # Values to build up the line chart.
    values = Reference(ws, min_col=content[0], min_row=content[1], max_col=content[2], max_row=content[3])
    # Names / categories for x-axis. In the example, the months.
    x_values = Reference(ws, range_string=range_axis)
    # Initialize LineChart object.
    chart = LineChart()
    # Add data to the LineChart object.
    chart.add_data(values, titles_from_data=True)
    # Set x-axis
    chart.set_categories(x_values)

    # Cosmetics for the graph.
    chart.title = 'Sales per month'
    chart.x_axis.title = 'Month'
    chart.y_axis.title = 'Ticket Sales (USD Mil)'
    chart.legend.position = 'b'
    ws.add_chart(chart, 'H1')

    wb.save(filename)
