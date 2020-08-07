from automagica import utilities


def excMarco(fileName, firstKey, secondKey):
    """
    打开excel文件，然后执行已经绑定组合键的宏,执行前要确认保存宏的文件已经打开
    :param fileName: 文件名
    :param firstKey: 组合键1
    :param secondKey: 组合键2
    :return: nothing
    """
    # 打开文件
    # 执行组合键
    # 保存文件
    # 关闭，退出
    pass


def excDirectoryMarco(path, date, firstKey, secondKey):
    """
    对文件夹里所有excel文件执行绑定组合键的宏，执行前要确认保存宏的文件已经打开
    :param path: 文件夹名
    :param date: 日期
    :param firstKey: 组合键1
    :param secondKey: 组合键2
    :return: nothing
    """
    # 以path下所有文件名生成一个list
    # 进行清洗，只保留excel文件
    # 进行清洗，只保留某日期生成的文件
    # 对list中每个文件执行excMarco(fileName,firstKey,secondKey)
    pass
