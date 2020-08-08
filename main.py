import datetime
from automagica.activities import Excel
from excMarco import excDirectoryMarco

if __name__ == '__main__':
    root = 'C:/Users/Administrator/PycharmProjects/taxFreeCustomerExcelFile/result'
    date = datetime.date.today().strftime("%Y-%m-%d")
    # date = '2020-07-31'
    currentDate = datetime.datetime.strptime(date, "%Y-%m-%d")
    macroFile = 'C:/Users/Administrator/AppData/Roaming/Microsoft/Excel/XLSTART/PERSONAL.XLSB'
    macroInfoSet = [
        ['中御', 'ctrl'],
        ['分界洲', 'ctrl'],
        ['大小洞天', ''],
        ['蜈支洲', ''],
        ['南山', ''],
        ['中御趣佰', ''],
        ['呀诺达', ''],
        ['玫瑰谷', ''],
        ['二次产出填报', ''],
        ['大树', ''],
        ['海棠湾', ''],
        ['海口免税店', '']
    ]
    excel = Excel(macroFile)

    for macroInfo in macroInfoSet:
        path = root + '/' + macroInfo[0]
        excDirectoryMarco(path, currentDate, macroInfo[1])

    excel.quit()
