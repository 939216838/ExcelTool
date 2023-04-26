# 这是一个示例 Python 脚本。
import time
import openpyxl
import win32con
import win32gui
from openpyxl.cell import Cell
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


# 按 Shift+F10 执行或将其替换为您的代码。
# 按 双击 Shift 在所有地方搜索类、文件、工具窗口、操作和设置。


def print_hi(name):
    # 在下面的代码行中使用断点来调试脚本。
    print(f'Hi, {name}')  # 按 Ctrl+F8 切换断点。


def tiaozheng_size():
    print("调整大小")
    window = win32gui.FindWindow(None, "AI问答聊")
    # 获得窗口句柄
    rect = win32gui.GetWindowRect(window)
    # 将窗口放到最前
    win32gui.SetForegroundWindow(window)
    if (win32gui.IsIconic(window)):
        win32gui.ShowWindow(window, win32con.SW_RESTORE)
    time.sleep(0.3)
    win32gui.SetWindowPos(window, win32con.HWND_TOPMOST, 50, 50, 400, 1000, win32con.SWP_SHOWWINDOW)
    time.sleep(0.3)


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    tiaozheng_size()


    pass
    # print_hi('PyCharm')
    # s1 = {"Google", "Runoob", "Taobao"}
    # # 给s1添加元素'baidu'
    # print(s1)
    # s1.add('baidu')
    # print(s1)
    #
    # # 给s1添加元素[1,2,3]
    # s1.update((1, 2, 3))
    # s1.update({'name': 'kate'})  # 对于字典，加入的是字典的key,不是value
    # # s1.update([1, 2, {3, 4}])  # TypeError: unhashable type: 'set'
    # print(s1)
    #
    # # 给s1添加元素[1,2,3]，{'bai'，'du'}
    # s1.update([1, 2, 3], {'bai', 'du'})
    # print(s1)
    # settlementInformationList = []
    # information = SettlementInformation("", "", "", "", "", "", "")
    # for i in range(10):
    #     information.name = i
    #     settlementInformationList.append(SettlementInformation(information.name,"", "", "", "", "", ""))
    #
    # for mation in settlementInformationList:
    #     print(mation)
    # name = "   单洪峰   "
    # strip1 = strip(name)
    # name_strip = name.strip()
    # print(strip1)
    # print(name_strip)
    # print(name)
    # string = 'asdr1234'
    #
    # print(string.isdigit())
    # string = '1234'
    #
    # print(string.isdigit())
    # num = 5521.88 + \
    #       4032.46 + \
    #       1140.19 + \
    #       2772.88 + \
    #       513.39 + \
    #       119.39 + \
    #       1295.4
    #
    # new_num = format(num, '.2f')
    # print(num)
    # print(new_num)
    #
    # f = 107920.00
    # values = format(f, '.2f')
    # rstrip = values.rstrip("0").rstrip(".")
    # print(rstrip)
    # month =4
    #
    # column = 2 + (month - 1) * 3
    # print(column)
    # # 列出文件夹下所有文件
    # python
    # 定义一个字典
    # dict = {'name': '张三', 'age': 20, 'gender': '男'}
    #
    # # 判断是否包含某个key
    # if 'name' in dict:
    #     print('字典中包含name这个key')
    #
    # if 'address' not in dict:
    #     print('字典中不包含address这个key')
    ###
    #
    # 输出结果为：
    #
    #
    # 字典中包含name这个key
    # 字典中不包含address这个key
    # #
    #
    # 也可以使用
    # `dict.get(key[, default])` 方法来获取字典中指定key的值，如果该key不存在则返回默认值（如果有指定的话）。
    #
    # 示例代码如下：

    # python
    # 定义一个字典
    # dict = {'name': '张三', 'age': 20, 'gender': '男'}
    #
    # # 获取指定key的值
    # name = dict.get('name')
    # print(name)
    #
    # # 当key不存在时，设定默认值
    # address = dict.get('address')
    # print(address)
    # print(dict)
    # #

    # 输出结果为：
    #
    # #
    # 张三
    # 未知
    # #
    # for i in range(1, 13):
    #     print(i)
    # ai_ = ["B", "E", "H", "K", "N", "Q", "T", "W", "Z", "AC", "AF", "AI"]
    # new_list = [x + '4' for x in ai_]
    # print(new_list)
    # join = "="+"+".join(new_list)
    # print(join)
    # list = []
    # list.append(3)
    # list.append(5)
    # list.append(4)
    # list.append(1)
    # list.append(2)
    # print(list)
    # list = ["延吉市延河水库有限公司",
    #         "延吉市五道水库有限公司",
    #         "延边桃源水电总厂",
    #         "延边汇茂能源开发有限公司",
    #         "延边东电茂霖水能奶头河发电有限公司",
    #         "五虎岭电站",
    #         "汪清县满台城综合开发有限公司",
    #         "汪清县荒坪洋水电站有限公司",
    #         "龙井市龙江水电站",
    #         "龙井市龙河水利水电开发有限公司",
    #         "龙井市豆满江水电有限公司",
    #         "两江电站",
    #         "吉林省珲春老龙口供水有限责任公司",
    #         "吉林省地方水电有限公司安图分公司",
    #         "珲春市华源水电投资开发有限公司",
    #         "珲春市华龙源水电有限责任公司",
    #         "和龙市新兴水力发电有限公司",
    #         "和龙市松月水力发电有限公司",
    #         "和龙市龙门水力发电有限公司",
    #         "国电电力发展股份有限公司磨盘山电站",
    #         "敦化市中瑞水电开发有限公司",
    #         "敦化市永兴丹江电站有限公司",
    #         "敦化市小石河二级电站",
    #         "敦化市祥源黑石三级电站有限公司",
    #         "敦化市沙河弘源水电梯级开发有限责任公司",
    #         "敦化市黑石梯级电站有限公司",
    #         "敦化市和鑫上石电站有限公司",
    #         "敦化茂霖水能上沟发电有限公司",
    #         "长白山保护开发区华龙水电有限公司",
    #         "安图县三零三电站有限公司",
    #         "安图县冰山水力发电有限公司",
    #         "安图天正光明发电有限公司",
    #         "安图长白山明月湖水资源开发有限公司"]
    # setMap = set(list)
    # print(len(list))
    # print(len(setMap))
    # for row in range(5, 100):
    #     print(row)
    #     thydropowerTotalDataList = {
    #                       '1':{'name':'Tom','score':90,'age':18},
    #                       '2':{'name':'Jerry','score':88,'age':17},
    #                       '王':{'name':'Kate','score':95,'age':19},
    #                       '4':{'name':'John','score':86,'age':16}
    #                       }
    #
    #
    #     col_e_data = ["5",'王', '1', '4', '2']
    #     # 使用字典推导式，将MonthEndReport.hydropowerTotalDataList按照col_e_data顺序排列
    #     hydropowerTotalDataList = {key: thydropowerTotalDataList.get(key,None) for key in col_e_data}
    #
    #     # 输出按照col_e_data顺序排列后的字典
    #     print(hydropowerTotalDataList)
    #
    #
