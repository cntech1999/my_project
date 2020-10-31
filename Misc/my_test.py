import xlrd
import operator
import locale

# For current year
file_path_current = './2019_June28.xlsx'

work_book = xlrd.open_workbook(file_path_current)
data_sheet = work_book.sheet_by_name('Sheet1')

all_dic_current = {}
div_dic_current = {}
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
for i in range(1,data_sheet.nrows):
    data_read = data_sheet.row_values(i)
    year = int(data_read[0][0:4])
    name = data_read[2]
    final_price = locale.atof(data_read[4])
    divident = data_read[9]
    if divident != '':
        divident = locale.atof(data_read[9])
        div_earning_ratio = float(data_read[10])
        all_dic_current[name] = {'year': year,'final_price': final_price, 'divident': divident, 'div_earning_ratio': div_earning_ratio}
        all_dic_current[name][year] = {'final_price': final_price, 'divident': divident, 'div_earning_ratio': div_earning_ratio}

        div_dic_current[name] = div_earning_ratio
        # MongoDB insert ...
        # print(name, all_dic[name])
        # print(name, all_dic[name]['divident'])

# name = '오리온'
# print(name, div_dic[name])

sorted_div_current = sorted(div_dic_current.items(), key=operator.itemgetter(1), reverse = True)
# print(file_path_current)
# for i in range(0, 5): #len(sorted_div)
#     print(sorted_div_current[i])
#
# tmp_name = sorted_div_current[0][0]
# print(tmp_name, all_dic_current[tmp_name]['final_price'])




# For next year
file_path_next = './2020_June30.xlsx'

work_book = xlrd.open_workbook(file_path_next)
data_sheet = work_book.sheet_by_name('Sheet1')

all_dic_next = {}
div_dic_next = {}
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
for i in range(1,data_sheet.nrows):
    data_read = data_sheet.row_values(i)
    year = int(data_read[0][0:4])
    name = data_read[2]
    final_price = locale.atof(data_read[4])
    divident = data_read[9]
    if divident != '':
        divident = locale.atof(data_read[9])
        div_earning_ratio = float(data_read[10])
        all_dic_next[name] = {'year': year,'final_price': final_price, 'divident': divident, 'div_earning_ratio': div_earning_ratio}
        div_dic_next[name] = div_earning_ratio
        # print(name, all_dic[name])
        # print(name, all_dic[name]['divident'])

# name = '오리온'
# print(name, div_dic[name])

sorted_div_next = sorted(div_dic_next.items(), key=operator.itemgetter(1), reverse = True)
# print(file_path_next)
# for i in range(0, len(sorted_div_next)): #len(sorted_div)
#     print(sorted_div_next[i])


sorted_div_final = {}
for i in range(0, int(len(sorted_div_current)*0.05)):
# for i in range(int(len(sorted_div_current) * 0.70), int(len(sorted_div_current) * 0.75)):
    name = sorted_div_current[i][0]
    # print(name)
    # if name in all_dic_next:
    #     print('True')
    # else:
    #     print(':P')

    # Error check if name is existing in the next year
    # print(sorted_div_next)
    if name in all_dic_next:
        growth_rate = (all_dic_next[name]['final_price'] - all_dic_current[name]['final_price']) / all_dic_current[name][
        'final_price'] * 100
        sorted_div_final[name] = {'div_earning_ratio': all_dic_current[name]['div_earning_ratio'], 'growth_rate': round(growth_rate, 2)}
        # print(name, all_dic_current[name]['div_earning_ratio'], round(growth_rate, 2))
        # print(name, sorted_div_final[name])

# print(len(sorted_div_final))
# print(sorted_div_final)

dic_keys = sorted_div_final.keys()
total_earning = 0
invest_per_stock = 1000000
for k in dic_keys:
    name = k
    # earning = int(invest_per_stock/all_dic_current[name]['final_price'])*all_dic_current[name]['final_price']*sorted_div_final[name]['growth_rate']*0.01
    # print(earning)
    # earning = invest_per_stock * sorted_div_final[name]['growth_rate']*0.01 + invest_per_stock * sorted_div_final[name]['div_earning_ratio']*0.01
    earning = invest_per_stock * sorted_div_final[name]['growth_rate'] * 0.01
    total_earning = total_earning + earning

invest = len(sorted_div_final)*invest_per_stock
earning_ratio = total_earning/invest * 100
remaining = invest + total_earning
print(invest, round(total_earning, 2))
print(round(earning_ratio),2)
print(remaining)



