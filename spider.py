from lxml import etree
import requests
import openpyxl
import re
import time

# 查找第三个td没有span标签的字符串起始，并进行拼接修改
def confirm_html(html):
    p = re.compile('<td STYLE="width:15%">\s{10,}</td>')
    while True:
        answer = p.search(html)
        if answer:
            mark = answer.span()
            html = html[:mark[0]] + '<td STYLE="width:15%"><span></span></td>' + html[mark[1]+1:]
        else:
            break
    return html

# 抓取页面html爬虫
def spider(url):
    # 设置headers及cookies
    headers = {'user-agent':'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'}
    cookies = {'JSESSIONID':'C82AC851C50F8EFB3D9F73F405A78369'}

    # 获取页面html并返回
    r = requests.get(url, headers=headers, cookies=cookies)
    result = r.text
    print('页面html获取完毕!')
    return result

# 获取html表单中的关键信息数据
def get_data(html):
    # html第一次过滤，找到表格div
    left = html.find(r'<div class="grid"')
    html = html[left:]

    # html第二次过滤，找到表格主体<tr>
    left = html.find('<tbody>')
    right = html.find('</tbody>')
    html = html[left+7:right]

    # 查找第三个td没有span标签的字符串起始，并进行拼接修改
    html = confirm_html(html)

    # 使用lxml获取表格关键信息并通过datas返回
    datas = etree.HTML(html).xpath('//td[1] | //td[2] | //td[3]/span[1] | //td[4] | //td[5] | //td[11]/a')
    print('当前页面数据获取完毕')
    return datas

# 将关键信息数据写入ex列表
def write_list(datas):
    print('开始写入当前页面数据')

    # 找到循环次数，数据总体量为6的倍数
    length = len(datas)
    times = length // 6

    # 开始循环写入数据，写入数据格式为[data1, data2, data3, data4, data5, data6, data7]
    for i in range(times):
        # 每次循环重置写入的list
        list2 = []
        for j in range(6):
            if datas[i*6+j].text == None:
                list2.append('该内容为空')
            else:
                list2.append(datas[i*6+j].text)
        url_address = 'https://m.simpletour.com/mobile/gateway/wechat/tourism/detail/' + list2[0]
        # 添加data7
        list2.append(url_address)
        # 将数据写入ex列表
        ex.append(list2)

    print('当前页面数据写入完毕')

# 将ex列表写入Excel
def write_excel(path, data_list):
    print('所有数据获取完毕，开始写入Excel')

    # 创建Excel工作簿
    wb = openpyxl.Workbook()
    # 激活Excel工作表
    sheet = wb.active
    # 工作表命名
    sheet.title = '车位'

    # 循环写入Excel数据
    for i in range(len(data_list)):
        for j in range(len(data_list[i])):
            sheet.cell(row=i + 1, column=j + 1, value=str(ex[i][j]))

    # 保存到指定路径下的Excel文件
    wb.save(path)

    print('Excel写入完毕，请到路径%s下查看文件!' %path)

# 定义保存到Excel的数据类别
ex = [['自营编号', '自营商品名', '目的地编号及商品名', '所属目的地', '所属线路', '状态', '微站链接']]
# 定义Excel文件保存路径，这里为当前路径下的"自营车位.xlsx"文件
path = '自营车位.xlsx'

if __name__ == "__main__":
    # 开始爬虫
    for i in range(1, 91):
        url = 'http://appadmin.simpletour.com/app/tourism/list?index='+str(i)+'&size=10&tag=SELF'
        print('开始爬取第%d页车位数据!' %i)
        html = spider(url)
        datas = get_data(html)
        write_list(datas)
        # 设置爬虫间隔，防止崩溃
        time.sleep(3)

    # 写入Excel
    write_excel(path, ex)