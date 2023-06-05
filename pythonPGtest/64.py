import requests
import json
import openpyxl

# 打开Excel文件
workbook = openpyxl.load_workbook('C:\\Users\\gaoxudong\\PycharmProjects\\TesTools\\data\\data.xlsx')

# 选择第一个工作表
worksheet = workbook.worksheets[0]

# 获取第一列数据
listhotel=[]
listyc=[]
listyc2=[]
for cell in worksheet['A']:
    listhotel.append(str(cell.value))
print(listhotel)
# print(len(listhotel))
# listhotel=["x100820"]
n=0
m=0

for i in listhotel:
    try:
        print(i)
        #获取静态：url
        staticurl="http://dashboard2.mis.elong.com/proxy/10.152.15.58:9192/global/globalqueryhotelroom"
        headers = {'Content-Type': 'application/json'}

        data={
            "hotelId": i,
            "batch": "2022-11-14 00:00:00",
            "supplierCode": 64,
            "syncType": 0,
            "contextId": "74abfab9-7c6a-4856-b034-bdedc68b6742",
            # "supplierName": "shengshengguoji",
            # "eanHC":True,
            "offset": 0
        }
        # 发送POST请求

        re = requests.post(url=staticurl, headers=headers, json=data)
        cre=re.json()
        print(cre)
        staic=cre['body']['hotelModelForCommon']
        key=cre['body']['key']
        redata = {
            "hotelModelForCommon": staic,
            "key": cre['body']['key'],
            "syncType": 0
        }
        print(redata)

        # 落地接口
        conteurl = "http://jhotelmaster.vip.elong.com:8080/api/saveOtaHotelInfo"
        contre = requests.post(url=conteurl, headers=headers, json=redata)
        print(contre.json())
        if contre.json()['retcode']==1001:
            listyc2.append(i)
            print(listyc2)

        n = n + 1
        print("执行次数：" + str(n))
    except Exception as e:
        print(e)
        m=m+1
        listyc.append(i)
        print(listyc)
        print("异常次数："+str(m))

