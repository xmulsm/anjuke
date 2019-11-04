import random
import urllib
from urllib import request
from lxml import etree
import xlwt
import time
import xlrd
from xlutils.copy import copy


def getanjuke(area, url):
    '''
    获取安居客二手房源数据
    :param area:地区
    :param url: 爬取的链接地址
    :return:
    '''
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36 SE 2.X MetaSr 1.0'
    # 传递给header，由于安居客有反爬机制，需要构造headers来骗过服务器
    headers = {'User-Agent': user_agent,
               'authority': 'xm.anjuke.com', 'scheme': 'https', 'upgrade-insecure-requests': '1',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
               'Accept-Language': 'zh-CN,zh;q=0.8', 'Referer': 'https://xm.anjuke.com/sale/?from=navigation', }
    try:
        oldwb = xlrd.open_workbook('安居客房源.xls')
        newwb = copy(oldwb)
        sheet = newwb.add_sheet(area)
        # 写入excel
        head = ["序号", "区域", "标题", "地址", "总价", "单价", "户型", "面积", "楼层", "建造年份", "经纪人", "链接"]
        # 写入首行
        for i in range(0, len(head)):
            sheet.write(0, i, head[i])
        newwb.save('安居客房源.xls')
    except FileNotFoundError:
        print('文件不存在')

    for index in range(0, 51):
        page_url = url + 'p' + str(index + 1)
        print('正在爬取第%d页数据，链接地址：%s' % (index, page_url))
        req = urllib.request.Request(page_url, headers=headers)
        response = urllib.request.urlopen(req)
        html = response.read().decode('utf-8')
        html_etree = etree.HTML(html)
        # print(html)
        title_list = html_etree.xpath('//*[@id="houselist-mod-new"]//a/@title')  # 获取标题
        address_list = html_etree.xpath('//*[@id="houselist-mod-new"]//span/@title')  # 获取地址
        totalprice_list = html_etree.xpath(
            '//*[@id="houselist-mod-new"]//span[@class="price-det"]/strong/text()')  # 获取总价
        unitprice_list = html_etree.xpath('//*[@id="houselist-mod-new"]//span[@class="unit-price"]/text()')  # 获取单价
        huxing_list = html_etree.xpath(
            '//*[@id="houselist-mod-new"]//div[@class="details-item"][1]/span[1]/text()')  # 获取户型
        size_list = html_etree.xpath(
            '//*[@id="houselist-mod-new"]//div[@class="details-item"][1]/span[2]/text()')  # 获取面积
        louceng_list = html_etree.xpath(
            '//*[@id="houselist-mod-new"]//div[@class="details-item"][1]/span[3]/text()')  # 获取楼层
        year_list = html_etree.xpath(
            '//*[@id="houselist-mod-new"]//div[@class="details-item"][1]/span[4]/text()')  # 获取建造年份
        brokername_list = html_etree.xpath('//*[@id="houselist-mod-new"]//span[@class="brokername"]/text()')
        url_list = html_etree.xpath('//*[@id="houselist-mod-new"]//a/@href')  # 获取房源链接

        # 写第一列
        for i in range(0, len(title_list)):
            sheet.write(i + 1 + index * len(title_list), 0, i + 1 + index * len(title_list))
            sheet.write(i + 1 + index * len(title_list), 1, area)
            sheet.write(i + 1 + index * len(title_list), 2, title_list[i])
            sheet.write(i + 1 + index * len(title_list), 3, address_list[i])
            sheet.write(i + 1 + index * len(title_list), 4, totalprice_list[i])
            sheet.write(i + 1 + index * len(title_list), 5, unitprice_list[i])
            sheet.write(i + 1 + index * len(title_list), 6, huxing_list[i])
            sheet.write(i + 1 + index * len(title_list), 7, size_list[i])
            sheet.write(i + 1 + index * len(title_list), 8, louceng_list[i])
            sheet.write(i + 1 + index * len(title_list), 9, year_list[i])
            sheet.write(i + 1 + index * len(title_list), 10, brokername_list[i])
            sheet.write(i + 1 + index * len(title_list), 11, url_list[i])

        newwb.save('安居客房源.xls')

    time.sleep(0.5)  # 休眠0.5秒，防止被服务器拒绝


if __name__ == '__main__':
    area_dict = {'思明区': 'https://xm.anjuke.com/sale/siming/', '湖里区': 'https://xm.anjuke.com/sale/huli/',
                 '海沧区': 'https://xm.anjuke.com/sale/haicang/', '集美区': 'https://xm.anjuke.com/sale/jimei/',
                 '同安区': 'https://xm.anjuke.com/sale/tongana/',
                 '翔安区': 'https://xm.anjuke.com/sale/xiangana/'}  # 六个区的链接地址
    for area, url in area_dict.items():
        getanjuke(area, url)
