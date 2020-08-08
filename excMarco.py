from automagica.activities import get_files_in_folder, Excel
import datetime


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


def choseFile(inputList, filterDate):
    """
    按日期筛选符合要求的文件，输出到列表
    :param inputList: 输入的文件列表
    :param filterDate: 日期 "%Y-%m-%d"
    :return: 符合日期要求的文件列表
    """
    import os
    outputList = []
    for file in inputList:
        statInfo = os.stat(file)
        createTime = statInfo.st_ctime
        createDayArray = datetime.datetime.fromtimestamp(createTime)
        createDay = createDayArray.strftime("%Y-%m-%d")
        if createDay == filterDate:
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
    # 生成path下所有excel文件的列表
    fileList = get_files_in_folder(pathName, extension='xlsx')
    # 进行清洗，只保留某日期生成的文件
    filteredFileList = choseFile(fileList, filterDate)
    # 对list中每个文件执行excMarco(fileName,marco)
    for file in filteredFileList:
        try:
            excMacro(file, macroName)
        except:
            pass


if __name__ == '__main__':
    path = 'C:/Users/Administrator/PycharmProjects/taxFreeCustomerExcelFile/result/南山'
    date = datetime.date.today().strftime("%Y-%m-%d")
    macroFile = 'C:/Users/Administrator/AppData/Roaming/Microsoft/Excel/XLSTART/PERSONAL.XLSB'
    excel = Excel(file_path=macroFile)
    excDirectoryMarco(path, date, 'PERSONAL.XLSB!南山')
    excel.quit()
