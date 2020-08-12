import datetime

from automagica.activities import get_files_in_folder, Excel


def excMacro(fileName, macroName):
    """
    打开excel文件，然后执行已经绑定组合键的宏,执行前要确认保存宏的文件已经打开
    :param macroName: 宏名称
    :param fileName: 文件名
    :return: nothing
    """
    # 打开文件
    targetFile = Excel(file_path=fileName)
    # 执行宏
    targetFile.run_macro(macroName)


def filterFile(inputList, filterDate):
    """
    按日期筛选符合要求的文件，输出到列表,将不符要求的excel文件以及日期不对的文件掉
    :param inputList: 输入的文件列表
    :param filterDate: 日期 "%Y-%m-%d"
    :return: 符合日期要求的文件列表
    """
    import os.path
    outputList = []
    for file in inputList:
        statInfo = os.stat(file)
        createTime = statInfo.st_ctime
        createDayArray = datetime.datetime.fromtimestamp(createTime)
        createDay = createDayArray.strftime("%Y-%m-%d")
        if createDay != filterDate:
            os.remove(file)
        if createDay == filterDate and file.endswith('xlsx'):
            outputList.append(file)

    return outputList


def excDirectoryMarco(pathName, filterDate, macroName):
    """
    对文件夹里所有excel文件执行绑定组合键的宏，执行前要确认保存宏的文件已经打开
    :param macroName: 宏名称
    :param pathName: 文件夹名
    :param filterDate: 日期 "%Y-%m-%d"
    :return: nothing
    """
    # 生成path下所有文件的列表
    fileList = get_files_in_folder(pathName)
    # 进行清洗，只保留某日期生成的文件,并且返回xlsx后缀文件列表
    filteredFileList = filterFile(fileList, filterDate)
    # 对list中每个文件执行excMarco(fileName,marco)
    for file in filteredFileList:
        try:
            excMacro(file, macroName)
        except:
            pass


def dailyTask(date=None):
    """
    每日给生成的报备excel文件套用宏
    """
    root = 'C:/Users/Administrator/PycharmProjects/taxFreeCustomerExcelFile/result'
    if date == None:
        date = datetime.date.today().strftime("%Y-%m-%d")
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
        ['海口免税店', '海口免税店'],
        ['千古情', '千古情'],
        ['椰田', '椰田']
    ]

    for macroInfo in macroInfoSet:
        macroName = 'PERSONAL.XLSB!{}'.format(macroInfo[1])
        excel = Excel(file_path=macroFile)
        excDirectoryMarco(root + '/' + macroInfo[0], date, macroName)
        excel.quit()


if __name__ == '__main__':
    path = 'C:/Users/Administrator/PycharmProjects/taxFreeCustomerExcelFile/result/南山'
    date = datetime.date.today().strftime("%Y-%m-%d")
    macroFile = 'C:/Users/Administrator/AppData/Roaming/Microsoft/Excel/XLSTART/PERSONAL.XLSB'
    excel = Excel(file_path=macroFile)
    excDirectoryMarco(path, date, 'PERSONAL.XLSB!南山')
    excel.quit()
