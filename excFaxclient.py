"""
准备工作
1.打开faxclient.exe  等待N秒
2.点击地接(L)
3.移动到计调业务，停留N秒
4.移动到团队销售单总表，点击，等待N秒
新增销售单
5.点新增，移动到新增团队销售单，点击，等待N秒
6.点击客户编码，输入编码，等待N秒
7.各种点击录入。包括客人名，手机，大人数，儿童数，陪同数，以上可以用tab键来切换
8.点击录入操作说明
9.切换到接送信息页面，等N秒
10.录入来程航班，出发地，目的地，出发时间，到达时间，备注，以上可以用tab切换
11.录入返程航班，出发地，目的地，出发时间，到达时间，备注，以上可以用tab切换
12.切换到客人名单，移动到第一个名字栏，录入名字，按两下tab，然后录入身份证号，如果有提示框跳出，按回车
13.点击鼠标，然后按一次下，录入名字，按两下tab，然后录入身份证号如果有提示框跳出，按回车
14.依次录完所有客人名单，第N个客人是用点击后按n-1次下，然后再录入名字，按两下tab，录入身份证号。如果有提示框跳出，按回车
15.按保存，出提示框架回车，然后关掉下单界面。
"""

from automagica import *


def ready():
    # 打开联创，等待15秒
    run(r'E:\旅游业务在线管理系统\faxclient.exe')
    wait(15)
    # 点击地接(L)
    click(x=207, y=37)
    wait(0.5)
    # 移动到计调业务，停留N秒
    move_mouse_to(x=235, y=108)
    wait(1)
    # 移动到团队销售单总表，点击，等待N秒
    click(x=444, y=188)
    wait(5)


class SaleLine:

    def __init__(self, info):
        # 点新增，移动到新增团队销售单，点击，等待N秒
        click(x=304, y=84)
        wait(0.5)
        click(x=314, y=107)
        wait(1)
        # basicInfo为一个列表，内为cusCode,cusName,cusTel,adult,child,guide
        self.basicInfo = info['basicInfo']
        # AmountInfo为一个列表，内为adultPrice,childPrice
        self.AmountInfo = info['AmountInfo']
        # OpInfo为一个字符串
        self.OpInfo = info['OpInfo']
        # AirInfo为一个列表，内为air1,from1,to1,departure1,arrive1,air2,from2,to3,departure4,arrive5
        self.AirInfo = info['AirInfo']
        # routeId为一个字符串，是线路编码
        self.routeId = info['routeId']
        # date是一个列表，格式为['2020','08','19','2020','08','23']
        self.date = info['date']
        # cusInfo为一个列表，内为客人名字，身份证号。格式为[(name1,id1),(name2,id2)....]
        self.cusInfo = info['cusInfo']
        self.cusNumber = len(self.cusInfo)

    def selectRoute(self):
        click(x=558, y=385)
        wait(1.5)
        click(x=252, y=203)
        typing(self.routeId)
        click(x=377, y=198)
        wait(0.5)
        double_click(x=493, y=250)
        wait(1)

    def inputBasicInfo(self):
        # 录入客户编码
        set_to_clipboard(self.basicInfo[0])
        click(x=314, y=205)
        press_key_combination('ctrl', 'v')
        # 录入客人名
        set_to_clipboard(self.basicInfo[1])
        click(x=303, y=273)
        press_key_combination('ctrl', 'v')
        # 录入客人电话
        set_to_clipboard(self.basicInfo[2])
        click(x=303, y=298)
        press_key_combination('ctrl', 'v')
        # 录入人数
        press_key('tab')
        typing(self.basicInfo[3])
        press_key('tab')
        typing(self.basicInfo[4])
        press_key('tab')
        typing(self.basicInfo[5])

    def inputAmount(self):
        click(x=855, y=234)
        typing(self.AmountInfo[0])
        click(x=941, y=234)
        typing(self.AmountInfo[1])
        click(x=1096, y=403)
        wait(0.5)
        click(x=1031, y=434)

    def inputOpDescription(self):
        set_to_clipboard(self.OpInfo)
        click(x=658, y=565)
        press_key_combination('ctrl', 'v')

    def inputAirInfo(self):
        click(x=313, y=487)
        wait(0.2)
        for i in range(10):
            if i == 0:
                set_to_clipboard(self.AirInfo[i])
                click(x=453, y=584)
                press_key_combination('ctrl', 'v')
                press_key('tab')
            elif i == 5:
                set_to_clipboard(self.AirInfo[i])
                click(x=453, y=601)
                press_key_combination('ctrl', 'v')
                press_key('tab')
            else:
                set_to_clipboard(self.AirInfo[i])
                press_key_combination('ctrl', 'v')
                press_key('tab')

    def inputCusInfo(self):
        click(x=420, y=488)
        for i in range(self.cusNumber):
            (name, idNumber) = self.cusInfo[i]
            click(x=334, y=537)
            if i > 0:
                for t in range(i):
                    press_key('pagedown')
            set_to_clipboard(name)
            press_key_combination('ctrl', 'v')
            press_key('tab')
            wait(0.5)
            press_key('tab')
            wait(0.5)
            press_key('tab')
            set_to_clipboard(idNumber)
            press_key_combination('ctrl', 'v')

    def inputDate(self):
        click(x=303, y=432)
        for i in range(6):
            if i == 2:
                typing(self.date[i])
                press_key('tab')
            elif i == 5:
                typing(self.date[i])
            else:
                typing(self.date[i])
                press_key('rightarrow')

    def save(self):
        click(x=245, y=150)


if __name__ == '__main__':
    ready()
    infor = {
        # basicInfo为一个列表，内为cusCode,cusName,cusTel,adult,child,guide
        'basicInfo': ['10017', '蒋烨1大1小13273275269已收1210元', '13917027355 15000608569', '1', '1', '0'],
        # AmountInfo为一个列表，内为adultPrice,childPrice
        'AmountInfo': ['600', '210'],
        # OpInfo为一个字符串
        'OpInfo': '安排1标，房差已补  4晚柏瑞/海湾维景 D线/5天4晚 小孩含餐车  不能强制推车购，注意方式\n携程网络主推产品客人，请导游务必注重服务，做好网络点评引导工作，引导满分好评，谢谢\n（车上配备百宝箱）',
        # AirInfo为一个列表，内为air1,from1,to1,departure1,arrive1,air2,from2,to3,departure4,arrive5
        'AirInfo': ['9c8801', '上海', '三亚', '18:20', '21:35', '9c8802', '三亚', '上海', '22:30', ''],
        # routeId为一个字符串，是线路编码
        'routeId': 'sy025',
        # cusInfo为一个列表，内为客人名字，身份证号。格式为[(name1,id1),(name2,id2)....]
        'cusInfo': [('蒋烨', '310107197909192417'), ('蒋宥希', '310107201302161728')],
        # date是一个列表，格式为['2020','08','19','2020','08','23']
        'date': ['2020', '08', '19', '2020', '08', '23']
    }
    saleLine = SaleLine(infor)
    saleLine.selectRoute()
    saleLine.inputBasicInfo()
    saleLine.inputDate()
    saleLine.inputAmount()
    saleLine.inputOpDescription()
    saleLine.inputAirInfo()
    saleLine.inputCusInfo()
