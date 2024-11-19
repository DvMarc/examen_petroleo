import xlwings as xw
import sys
import os

project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.append(project_root)
import models.caudal as caudal

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    PWF =  'Pwf'
    qmax = 250
    pr = 2400

    valuesPwf = sheet.range(PWF).value
    valoresCaudal = []
    for pwf in valuesPwf:
        q0 = caudal.caudal(qmax, pr, pwf)
        print(q0)
        valoresCaudal.append(q0)

    print(valoresCaudal)

    chart1 = sheet.charts.add()
    chart1.name = "Pwf vs Q0"
    chart1.chart_type = "xy_scatter_lines"

    cell_position_1 = sheet.range("D1")
    chart1.left = cell_position_1.left
    chart1.top = cell_position_1.top
    chart1.width = 400
    chart1.height = 300

    chart1.api[1].SeriesCollection().NewSeries()
    chart1.api[1].SeriesCollection(1).XValues = valoresCaudal
    chart1.api[1].SeriesCollection(1).Values = valuesPwf
    chart1.api[1].SeriesCollection(1).Name = "Relación Pwf vs Q0"

    chart1.api[1].ChartTitle.Text = "Relación Pwf vs Q0"
    chart1.api[1].Axes(1).HasTitle = True
    chart1.api[1].Axes(1).AxisTitle.Text = "Q0 (stb/d)"
    chart1.api[1].Axes(2).HasTitle = True
    chart1.api[1].Axes(2).AxisTitle.Text = "Pwf (psi)"

@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("poes.xlsm").set_mock_caller()
    main()
