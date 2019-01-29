import re
import openpyxl
from wxpy import *


bot = Bot(console_qr=True, cache_path=True)
order = bot.groups().search(r'订单群')
province = ['北京', '天津', '上海', '重庆', '河北', '山西', '辽宁', '吉林', '黑龙江', '江苏', '浙江', '安徽',
            '福建', '江西', '山东', '河南', '湖北', '湖南', '广东', '海南', '四川', '贵州', '云南', '陕西',
            '甘肃', '青海', '台湾', '内蒙古自治区', '广西壮族自治区', '西藏自治区', '宁夏回族自治区',
            '新疆维吾尔自治区', '香港特别行政区', '澳门特别行政区']
products = [""]

@bot.register(order, except_self=False)
def reply_order(msg):
    cont = msg.text
    if '数量' in cont:
        wb = openpyxl.load_workbook(filename="order_book.xlsx")
        ws = wb.active
        rows = ws.max_row
        return '订单共计{}条!'.format(rows-1)
    if '订单' in cont:
        wb = openpyxl.load_workbook(filename="order_book.xlsx")
        ws = wb.active
        rows = tuple(ws.rows)[1:]
        order = ""
        for row in rows:
            print(row)
            info = [cell.value for cell in row if cell.value]
            print(info)
            order += ",".join(info) + '\n'
            print(order)
        return order

    print(cont)
    # 将文本中的“，”替换为空格,方便查看字符边缘
    cont = cont.replace('，', ' ')
    # 首先找到电话
    find = re.search(r'\d{11}', cont)
    if find is None:
        return "没有找到11位的手机号码!"
    tel = find.group()

    # 将消息中的电话删除后，再过滤姓名和地址
    cont = cont.replace(tel, '')
    name, addr = None, None

    finds = re.findall(r"\b\w+\b", cont)
    for i in finds:
        size = len(i)
        if size <= 8:
            # 如果文字数量小于8，则视为名称
            name = i
        else:
            # 对大于8个的文本，看看是否包含省份
            for each in province:
                if each in i:
                    addr = i
                    break
    if name is None:
        return "没有找到小于8个汉字名字,!"
    if addr is None:
        return "没有找到包含省、市、直辖市、特别行政区的地址!"
    # 买方、卖方、联系人、电话、产品名称、规格型号、数量、单价、总价（是否包含运费）、交付方式、
    print("日期：{} 姓名：{} 电话：{} 地址：{}".format(msg.create_time, name, tel, addr))

    date = msg.create_time.strftime('%Y-%m-%d %H:%M:%S')
    prod = ''
    price = ""
    number = ""
    # 日期 商品 价格 数量 买家 电话 地址
    order_data = [prod, price, number, name, tel, addr, date]
    # print(order_data)
    wb = openpyxl.load_workbook(filename="order_book.xlsx")
    ws = wb.active
    ws.insert_rows(2)
    for clo in range(1, 8):
        ws.cell(row=2, column=clo).value = order_data[clo-1]
    # 文件打开时，保存失败！无报错!
    wb.save('order_book.xlsx')

    return "总共添加{}条订单!".format(ws.max_row-1)

embed()


