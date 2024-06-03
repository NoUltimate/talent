from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment

# 打开Excel文件
workbook = openpyxl.load_workbook(
    'C:/Users/Administrator/PycharmProjects/talent/house/共有产权房人员清单2023.8.10(1).xlsx')
# 选择第一个工作表
sheet = workbook["Sheet1"]
for (i, row1) in enumerate(sheet.iter_rows()):
    month = "1"
    day = "1"
    if i == 0: continue
    value1 = str(sheet["H" + str(i + 1)].value)
    if value1 is None or value1 == "None":
        continue
    print(value1)
    year = value1[0:4]
    if len(value1) > 4:
        month = value1[4:6]
        if month[0] == '0':
            month = month[1:]
    if len(value1) > 6:
        day = value1[6:]
    sheet["H" + str(i + 1)].value = datetime.strptime(year+"/"+month+"/"+day, "%Y/%m/%d")
workbook.save('共有产权房人员清单.xlsx')
