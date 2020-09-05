import sseclient
import json
import datetime
import xlsxwriter


#股票代码
stockode='002425'
#secids
secids='0.'+str(stockode)

declareData={
    "f12":"代码",
    "f14":"名称",
    "f100":"所属板块",
    "f102":"所属地区板块",
    "f103":"所属概念板块",
    "f2":"最新价",
    "f3":"涨跌幅",
    "f4":"涨跌额",
    "f5":"总手",
    "f30":"现手",
    "f31":"买入价",
    "f32":"卖出价",
    "f18":"昨收",
    "f6":"成交额",
    "f8":"换手率",
    "f7":"振幅",
    "f10":"量比",
    "f22":"涨速",
    "f9":"市盈率",
    "f15":"最高价",
    "f16":"最低价",
    "f17":"开盘价",
    "f62":"主力净流入",
    "f63":"集合竞价",
    "f64":"超大单流入",
    "f65":"超大单流出",
    "f66":"超大单净额",
    "f69":"超大单净占比",
    "f70":"大单流入",
    "f71":"大单流出",
    "f72":"大单净额",
    "f75":"大单净占比",
    "f76":"中单流入",
    "f77":"中单流出",
    "f78":"中单净额",
    "f81":"中单净占比",
    "f82":"小单流入",
    "f83":"小单流出",
    "f84":"小单净额",
    "f87":"小单净占比",
    "f88":"当日DDX",
    "f89":"当日DDY",
    "f90":"当日DDZ",
    "f91":"5日DDX",
    "f92":"5日DDY",
    "f94":"10日DDX",
    "f95":"10日DDY",
    "f97":"DDX飘红天数(连续)",
    "f98":"DDX飘红天数(5日)",
    "f99":"DDX飘红天数(10日)",
    "f38":"总股本",
    "f39":"流通股",
    "f36":"人均持股数",
    "f112":"每股收益",
    "f113":"每股净资产",
    "f37":"净资产收益率(加权)",
    "f40":"营业收入",
    "f41":"营业收入同比",
    "f42":"营业利润",
    "f43":"投资收益",
    "f44":"利润总额",
    "f45":"净利润",
    "f46":"净利润同比",
    "f47":"未分配利润",
    "f48":"每股未分配利润",
    "f49":"毛利率",
    "f50":"总资产",
    "f51":"流动资产",
    "f52":"固定资产",
    "f53":"无形资产",
    "f54":"总负债",
    "f55":"流动负债",
    "f56":"长期负债",
    "f57":"资产负债比率",
    "f58":"股东权益",
    "f59":"股东权益比",
    "f60":"公积金",
    "f61":"每股公积金",
    "f26":"上市日期"


}
#获取key数组
keyArr=[]
#获取Value数组
valueArr=[]
#获取需要查询的字段
feilds=""
for key in declareData:
    feilds=feilds+key+","
    keyArr.append(key)
    valueArr.append(declareData[key])



#返回需要显示的列
feilds=feilds[:-1]

url="https://62.push2.eastmoney.com/api/qt/ulist/sse?invt=3&pi=0&pz=1&mpi=2000&secids="+secids+"&ut=6d2ffaa6a585d612eda28417681d58fb&fields="+feilds



#申明预期插入excel 数组
excelData=[[0,0] for i in range(len(keyArr))]


def with_urllib3(url):
    import urllib3
    http = urllib3.PoolManager()
    return http.request('GET', url, preload_content=False)

#获得需要返回的数据
def getTargetData():
    data = {}
    response = with_urllib3(url)  # or with_requests(url)
    client = sseclient.SSEClient(response)
    for msg in client.events():
        if msg.data != None:
            data = json.loads(msg.data)["data"]["diff"]["0"]
        else:
            data = None
        break
    client.close()
    return data

#处理数据，将数据转为万、亿
def str_of_num(num):
    '''
    递归实现，精确为最大单位值 + 小数点后2位
    '''
    oldnum=num
    num=abs(num)
    def strofsize(num, level):
        if level >= 2:
            return num, level
        elif num >= 10000:
            num /= 10000
            level += 1
            return strofsize(num, level)
        else:
            return num, level
    units = ['', '万', '亿']
    num, level = strofsize(num, 0)
    if level > len(units):
        level -= 1

    if oldnum>0:
       return '{}{}'.format(round(num, 2), units[level])
    else:
       return '{}{}'.format(round(-num, 2), units[level])

#处理数据格式
def handData(key,value):
    val=''

    if key in ['f2','f4','f31','f18','f9','f15','f16','f17','f90'] :
        val=value/100
    elif key in ['f3','f7','f8','f22','f69','f75','f81','f87']:
        val=str(value/100)+'%'
    elif key in ['f5','f6','f62','f64','f65','f66','f70','f71','f72','f76','f77','f78','f82','f83',
                 'f84','f38','f39','f36','f40','f42','f43','f44','f45','f47','f50','f51','f52','f53','f54','f55','f56','f58','f60']:
        val=str_of_num(value)
    elif key in['f88','f89','f91','f92','f93','f94','f95']:
        val=value/1000
    elif key in['f63']:
        val=round(value/10000)
    elif key in ['f112','f113','f41','f46','f48','f49','f57','f59','f61']:
        val = round(value,2)
    elif key=='f10':
        val=value/100


    else:
        if value==None or value=='':
            val="-"
        else:
            val=value
    return val

#处理组装数据
def  handleExcel(data):
    for key in data:
        index=keyArr.index(key)
        val=handData(key,data[key])
        nodeArr=[]
        nodeArr.append(valueArr[index])
        nodeArr.append(val)
        excelData[keyArr.index(key)]=nodeArr
    return excelData

#数据写入excel中
def writeData(excelData):
    #获取当天的时间
    today=datetime.date.today()
    month=today.month
    todayStr=today.strftime("%Y-%m-%d")
    # 文件路径
    # filePath = './createExcel/'+todayStr+'.xlsx'
    filePath = str(stockode)+"_"+todayStr+'.xlsx'
    # 创建一个新的 Excel 文件，并添加一个工作表
    workbook = xlsxwriter.Workbook(filePath)
    # 定义单元格样式
    red_color = workbook.add_format({'color': 'black', 'bold': True, 'fg_color': 'red', 'border': 1})
    green_color = workbook.add_format({'color': 'black', 'bold': True, 'fg_color': 'green', 'border': 1})
    #添加一个Sheet,名称为 X天数据
    sheetName=todayStr+'的数据'

    worksheet = workbook.add_worksheet(sheetName)
    # 设置第一列(A) 单元格宽度为 20
    worksheet.set_column('A:BZ', 20)
    #解析数据，并写入
    for i in range(len(excelData)):
        worksheet.write(0,i,excelData[i][0])
        worksheet.write(1,i,excelData[i][1])
    # 关闭 Excel 文件
    workbook.close()

#数据

if __name__ == '__main__':
    data=getTargetData()
    print(data)
    excelData=handleExcel(data)
    print(excelData)
    print(1)
    writeData(excelData)

