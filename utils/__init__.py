def convert_to_snake_case(s):
    snake_case = ''
    for i, c in enumerate(s):
        if c.isupper() and i > 0:
            snake_case += '_'
        snake_case += c.lower()
    return snake_case
# 打包命令
#  pyinstaller --name ExcelTool --onefile --hidden-import=openpyxl --hidden-import=wx --add-data="./MonthEndReport/MonthEndReport.py;./MonthEndReport/" --add-data="./MonthEndReport/SolarPower.py;./MonthEndReport/" ./MainWindow/MainWindow.py
if __name__ == '__main__':
    print(convert_to_snake_case("fullOnlineNonNaturalPersonList"))