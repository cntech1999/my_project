import xlrd
import operator
import locale
from pymongo import MongoClient

file_2b_imported = ['2016.xlsx', '2017.xlsx', '2018.xlsx', '2019.xlsx', '2020.xlsx']
# print(file_list_2b_imported)
file_path = file_2b_imported[0]
# print(file_path_current)
work_book = xlrd.open_workbook(file_path)
data_sheet = work_book.sheet_by_name('Sheet1')

dic = {}
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
for i in range(1,data_sheet.nrows):
    data_read = data_sheet.row_values(i)
    year = int(data_read[0][0:4])
    code_id = data_read[1]
    name = data_read[2]
    final_price = locale.atof(data_read[4])
    per = data_read[6]
    pbr = data_read[8]
    div = data_read[9]
    div_earn_ratio = data_read[10]
    dic[name][year] = {'code_id': code_id, 'final_price': final_price, 'divident': div, 'div_earning_ratio': div_earn_ratio}
