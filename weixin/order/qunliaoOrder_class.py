import pymongo
import json
json.loads()
class Product(object):
    def __init__(self, name, price, number):
        self.name = name
        self.price = price
        self.number = number

    def __str__(self):
        return self.name

    def __repr__(self):
        return "商品：{}".format(self.name)


class Order(object):
    def __init__(self, client, tel, addr ,products, date, account):
        self.client = client
        self.tel = tel
        self.addr = addr
        self.products = products
        self.date = date
        self.account = account

    # def __str__(self):
    #     return '< 订单:{}  提交账号:{} >'.format(self.client, self.account)

    def get_info(self):
        prod = []
        pr_nb = []
        for product in self.products:
            prod.append(product.name)
            pr_nb.append(str(product.price) + 'x' + str(product.number))
        prod = '、'.join(prod)
        pr_nb = "、".join(pr_nb)
        return prod + ',' + pr_nb + ',' + self.client + ',' + self.tel + ',' + self.addr

    def save_excel(self, file):
        pass

text = '小粉刷，150x3，冯颖，18201098888，山东德州德城区'
text = text.replace(' ', '')
text = text.replace('，', ',')
finds = text.split(",")

order_info = finds
# 以顿号分离产品和价格
prods = order_info[0].split("、")
temps = order_info[1].split("、")
for i in range(len(prods)):
    price, number = None, None
    prod = prods[i]
    pr_nb = temps[i].replace('X', 'x').split("x")

    if len(pr_nb) == 1 and pr_nb[0].isdigit():
        price, number = pr_nb[0], '1'
    elif len(pr_nb) == 2 and pr_nb[0].isdigit() and pr_nb[1].isdigit():
        price, number = pr_nb[0], pr_nb[1]
    # else:
    #     return "错误：\n价格和数量格式不对！例:价格x数量(150x3)、价格(150)"
    prods[i] = Product(prod, int(price), int(number))
client = order_info[2]
tel = order_info[3]
# if len(tel) != 11 or tel.isdigit() is False:
#     return "错误：\n手机号码不是11位数或者不是纯数字号码！"
pro = ['北京', '天津', '上海', '重庆', '河北', '山西', '辽宁', '吉林', '黑龙江', '江苏', '浙江', '安徽',
       '福建', '江西', '山东', '河南', '湖北', '湖南', '广东', '海南', '四川', '贵州', '云南', '陕西',
       '甘肃', '青海', '台湾', '内蒙古', '广西壮族', '西藏', '宁夏回族', '新疆维吾尔', '香港', '澳门']
addr = None
for each in pro:
    if each in order_info[4]:
        addr = order_info[4]
        break
# if addr is None:
#     # return "错误：\n地址中没有包含省、直辖市、自治区、特别行政区的地址!"
date = '246554'
account = 'adf[p'

order = Order(client, tel, addr, prods, date, account)
print(order.get_info())
