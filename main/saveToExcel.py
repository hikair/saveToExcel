import datetime

import xlrd
import xlwt
import json


def dataCheck(sheet):
    rows = sheet.nrows
    assert rows > 2
    assert sheet.row(0)[0].value == 'key'
    assert sheet.row(0)[1].value == 'value'
    assert sheet.row(1)[0].value == 'json_path'
    assert sheet.row(2)[0].value == 'target_path'


# 保存至excel
def saveToExcel(content):
    json_data = json.loads(content)
    list_key = findFirstListKey(json_data)
    assert list_key is not None
    list_data = json_data[list_key]
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
    headers = getHeaders(list_data)
    # 写入头
    for i, j in enumerate(headers):
        sheet.write(0, i, j)

    size = len(list_data)
    for i in range(size):
        data = list_data[i]
        for key in data.keys():
            index = headers.index(key)
            if index >= 0:
                sheet.write(i + 1, index, str(data[key]))
    return book


# 获取表头
# 列表中的元素可能，属性数量不一致，取数量最多的
def getHeaders(list_data):
    index = 0
    attrs = 0
    for i, j in enumerate(list_data):
        t = len(j.keys())
        if t > attrs:
            index = i
            attrs = t
    return list(list_data[index].keys())


# 找到第一个数据类型为列表的key
def findFirstListKey(json_data):
    keys = list(json_data.keys())
    for key in keys:
        data_list = json_data[key]
        type_a = type(data_list)
        if type_a == list:
            return key
    return None


if __name__ == '__main__':
    file = 'resource/cmd.xls'
    wb = xlrd.open_workbook(filename=file)
    sheet1 = wb.sheet_by_index(0)
    # 校验数据
    dataCheck(sheet1)
    rows = sheet1.nrows
    json_path = sheet1.row(1)[1].value
    target_path = sheet1.row(2)[1].value
    if not target_path.endswith('.xls') and not target_path.endswith('.xlsx'):
        if not target_path.endswith('/'):
            target_path += '/'
        target_path = '{}{}.xls'.format(target_path, str(datetime.datetime.now()))

    with open(json_path) as f:
        content = f.read()
        book = saveToExcel(content)
        book.save(target_path)
