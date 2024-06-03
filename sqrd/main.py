import openpyxl
from openpyxl.styles import Alignment

# 打开Excel文件
workbook = openpyxl.load_workbook('C:/Users/Administrator/PycharmProjects/pythonProject/第一批授权认定用人单位名单(1).xlsx')
# 选择第一个工作表
sheet = workbook["汇总"]