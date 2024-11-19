import xlwings as xw


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    PWF =  'Pwf'

    valuesPwf = sheet[PWF]
    print(valuesPwf.value)

@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("poes.xlsm").set_mock_caller()
    main()
