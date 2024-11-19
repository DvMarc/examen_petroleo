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

@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("poes.xlsm").set_mock_caller()
    main()
