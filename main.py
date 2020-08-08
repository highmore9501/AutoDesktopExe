import datetime
from automagica.activities import Excel
from excMarco import excDirectoryMarco

if __name__ == '__main__':
    root = 'C:/Users/Administrator/PycharmProjects/taxFreeCustomerExcelFile/result'
    date = datetime.date.today().strftime("%Y-%m-%d")
    # date = '2020-07-31'
    macroFile = 'C:/Users/Administrator/AppData/Roaming/Microsoft/Excel/XLSTART/PERSONAL.XLSB'
    macroInfoSet = [
        ['中御', '中御'],
        ['分界洲', '分界洲'],
        ['大小洞天', '洞天奇遇'],
        ['蜈支洲', '蜈支洲'],
        ['南山', '南山'],
        ['呀诺达', '分界洲'],
        ['玫瑰谷', '玫瑰谷'],
        ['二次产出填报', '二次产出'],
        ['海棠湾', '免税店格式'],
        ['海口免税店', '海口免税店']
    ]

    for macroInfo in macroInfoSet:
        macroName = 'PERSONAL.XLSB!{}'.format(macroInfo[1])
        excel = Excel(file_path=macroFile)
        excDirectoryMarco(root+'/'+macroInfo[0], date, macroName)
        excel.quit()

