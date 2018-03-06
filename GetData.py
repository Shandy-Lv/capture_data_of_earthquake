# 导入模块
import re
import time
import urllib.request
import bs4
import xlsxwriter


def get_days(pagesoup):
    # 获取地震时间，先转换成字符串，由于地震地点不同字符串长度不同，但是时间符合固定长度。
    # 先翻转字符串，再取固定长度，再翻转获取时间
    # 获取的原始time示例：
    # ['origTime("8月11日19时26分四川阿坝州九寨沟县发生3.4级地震","2017-08-11 19:34:14");\n']
    # ['origTime("8月11日13时28分菲律宾发生6.3级地震","2017-08-11 14:04:15");\n']
    days = pagesoup.find_all(string=re.compile("\"[0-9]*-[0-9]*-[0-9]*"))
    days = '\n'.join(days)
    days = days[::-1]
    days = days[4:23]
    days = days[::-1]
    year = days[0:4]
    day = pagesoup.find_all(string=re.compile("[0-9]*月[0-9]*日[0-9]*时[0-9]*分"))
    day = "\n".join(day)
    end = day.index("分")
    day = day[0:end+1]
    days = year + "年" +day
    return days


def get_latitude(pagesoup):
    # 获取经度
    latitude = pagesoup.find_all(string=re.compile("subStringLocationLatitude\(\"[+-]*[0-9]*\.[0-9]*\"\)"))
    latitude = '\n'.join(latitude)
    latitudelist = latitude.split("\"")
    if len(latitudelist) > 1:
        latitude = latitudelist[1]
    else:
        latitude = ''
    return latitude


def get_longitude(pagesoup):
    longitude = pagesoup.find_all(string=re.compile("subStringLocationLongitude\(\"[+-]*[0-9]*\.[0-9]*\"\)"))
    longitude = '\n'.join(longitude)
    longitudelist = longitude.split("\"")
    if len(longitudelist) > 1:
        longitude = longitudelist[1]
    else:
        longitude = ''
    return longitude


def get_magnitude(pagesoup):
    # 获取地震等级
    magnitude = pagesoup.find_all(string=re.compile("级地震，震源深度"))
    magnitude = "\n".join(magnitude)
    start = magnitude.index("生")
    end = magnitude.index("级")
    level = magnitude[start + 1:end]
    # print(level)
    return level


def get_dept(pagesoup):
    # 获取深度
    deep = pagesoup.find_all(string=re.compile("shengdu\(\"[+-]*[0-9]*\.*[0-9]*\"\)"))
    deep = "\n".join(deep)
    deeplist = deep.split("\"")
    deep = deeplist[1]
    # print(deep)
    return deep


def write_into_excel(worksheet, row, days, latitude, longitude, level, deep):
    # Write into excel
    worksheet.write(row, 0, days)
    worksheet.write(row, 1, latitude)
    worksheet.write(row, 2, longitude)
    worksheet.write(row, 3, level)
    worksheet.write(row, 4, deep)
    print(days)


def get_data_and_write(pagesoup, worksheet, startrow):
    quake_list = pagesoup.find_all(href=re.compile("publish/dizhenj/464/479/20"))
    for t in quake_list:
        try:
            link = t.get("href")
            link = "http://www.cea.gov.cn"+link
            shake = urllib.request.urlopen(link)
            content = shake.read()
            # 获取编码
            # dammit = UnicodeDammit(content)
            # print(dammit.unicode_markup)
            # print(dammit.original_encoding)
            sp = bs4.BeautifulSoup(content, "html.parser", from_encoding="utf-8")
            time = get_days(sp)
            latitude = get_latitude(sp)
            longitude = get_longitude(sp)
            level = get_magnitude(sp)
            deep = get_dept(sp)
            write_into_excel(worksheet, startrow, time, latitude, longitude, level, deep)
            startrow += 1
            shake.close()
        except Exception():
            print(sp)
    return startrow


response = urllib.request.urlopen('http://www.cea.gov.cn/publish/dizhenj/464/479/index.html')
html = response.read()
# print(html)
time.sleep(1)
soup = bs4.BeautifulSoup(html, "html.parser")

workbook = xlsxwriter.Workbook('Data.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', "Time")
worksheet.write('B1', "Latitude")
worksheet.write('C1', "Longitude")
worksheet.write('D1', "Magnitude")
worksheet.write('E1', "Focal Depth（km）")

startRow = 1
row = get_data_and_write(soup, worksheet, startRow)
response.close()
page = soup.find_all(string=re.compile("共[0-9]*页"))
page = "\n".join(page)
temp_position = page.index("共")+1
page = page[temp_position:]
start_position = page.index("共")+1
end_position = page.index("页")
total_pages = page[start_position:end_position]
num = 2
for num in range(2, int(total_pages)):
    try:
        url = "http://www.cea.gov.cn/publish/dizhenj/464/479/index_"+str(num)+".html"
        print(num)
        response = urllib.request.urlopen(url)
        html = response.read()
        time.sleep(1)
        soup = bs4.BeautifulSoup(html, "html.parser")
        row = get_data_and_write(soup, worksheet, row)
        response.close()
    except Exception():
        print(url)
workbook.close()





