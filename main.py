import xlrd
import copy
import xlwt
import sort_function

file_name_0 = 'tables_fastest_0'
file_path_0 = '/home/yma/DHC/fastest_table' + '//' + file_name_0 + '.xls'
workbook_0 = xlrd.open_workbook(file_path_0)
sheets_0 = workbook_0.sheet_names()
worksheet_0 = workbook_0.sheet_by_name(sheets_0[0])
dir_dic_0 = {}
column_dir = 0
for i in range(worksheet_0.nrows):
    if worksheet_0.cell_value(i, column_dir) != '':
        dir_dic_0[worksheet_0.cell_value(i, column_dir)] = []
        for j in range(1, worksheet_0.ncols, 3):
            if worksheet_0.cell_value(i, j) != '':
                dir_dic_0[worksheet_0.cell_value(i, column_dir)].append(
                    [worksheet_0.cell_value(i, j), worksheet_0.cell_value(i, j + 1),
                     worksheet_0.cell_value(i, j + 2)])

file_name_1 = 'tables_fastest_1'
file_path_1 = '/home/yma/DHC/fastest_table' + '//' + file_name_1 + '.xls'
workbook_1 = xlrd.open_workbook(file_path_1)
sheets_1 = workbook_1.sheet_names()
worksheet_1 = workbook_1.sheet_by_name(sheets_1[0])
dir_dic_1 = {}
column_dir = 0
for i in range(worksheet_1.nrows):
    if worksheet_1.cell_value(i, column_dir) != '':
        dir_dic_1[worksheet_1.cell_value(i, column_dir)] = []
        for j in range(1, worksheet_1.ncols, 3):
            if worksheet_1.cell_value(i, j) != '':
                dir_dic_1[worksheet_1.cell_value(i, column_dir)].append(
                    [worksheet_1.cell_value(i, j), worksheet_1.cell_value(i, j + 1),
                     worksheet_1.cell_value(i, j + 2)])

dir_dic = {}
for i in dir_dic_0:
    dir_dic[i] = []
    for j in dir_dic_0[i]:
        if j not in dir_dic[i]:
            dir_dic[i].append(j)
    for k in dir_dic_1[i]:
        flag = False
        list_d = []
        for l in copy.deepcopy(dir_dic[i]):
            if l[0] == k[0]:
                flag = True
                list_d.append(l)
        if not flag:
            dir_dic[i].append(k)
        else:
            for m in list_d:
                e_prime = m[2]
                if k[2] < e_prime:
                    dir_dic[i].remove(m)
                    dir_dic[i].append(k)
        for entry in copy.deepcopy(dir_dic):
            if (k[0] < entry[0]) and (k[2] <= entry[2]):
                dir_dic[i].remove(entry)
    dir_dic[i].sort(key=sort_function.take_first, reverse=True)

file_name = 'tables_fastest'
file_path = '/home/yma/DHC/fastest_table' + '//' + file_name + '.xls'
sheet_name = 'TABLES_fastest'
workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)  # 新建一个工作簿
sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)  # 在工作簿中新建一个表格
row = 0
for i in dir_dic:
    sheet.write(row, 0, i)
    for j in range(len(dir_dic[i])):
        sheet.write(row, int(3 * j + 1), dir_dic[i][j][0])
        sheet.write(row, int(3 * j + 2), dir_dic[i][j][1])
        sheet.write(row, int(3 * j + 3), dir_dic[i][j][2])
    row += 1
workbook.save(file_path)
