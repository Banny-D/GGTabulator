from pandas import read_excel
from pandas import DataFrame
from pandas import ExcelWriter
from pandas import isna

# import openpyxl
# import pandas as pd

# 导入表格
try:
    table_data = read_excel('input.xlsx', header=None)
except FileNotFoundError:
    print('未找到文件\'input.xlsx\'')
    input('按任意键退出程序')
    exit()

empty_rows = table_data[table_data.isnull().all(axis=1)].index
# result = table_data.iloc[empty_rows + 1, 0].tolist()
# table_data.head()
# 均价
price_average = table_data.iloc[0,3]
print('均价：' + str(price_average))
# 导入简称表
try:
    symbol_data = read_excel('symbol.xlsx', header=None)
    # 将symbol表转换为字典，第1列为键，第2列为值
    symbol_dict = dict(zip(symbol_data.iloc[:, 0], symbol_data.iloc[:, 1]))
    # 在表中商品名称替换为简称，找不到对应的就不替换
    table_data.iloc[:, 0] = table_data.iloc[:, 0].apply(
        lambda x: symbol_dict.get(x, x))
except FileNotFoundError:
    print('未找到文件\'symbol.xlsx\'用于替换简称，程序继续运行')



# 表格行列数
table_rows = table_data.shape[0]
table_cols = table_data.shape[1]
# 找出分盒的开始位置和结束位置
group_rows_start = table_data[table_data.iloc[:,1].isnull()].index
group_rows_start = [0] + group_rows_start.tolist()
group_rows_start = [i+1 for i in group_rows_start]
group_rows_end = group_rows_start[1:]
group_rows_end = [i-1 for i in group_rows_end]
group_rows_end = group_rows_end + [table_rows]
# 分盒的数量
num_set_rows = len(group_rows_start)

# print(num_set_rows)
# print(group_rows_start)
# print(group_rows_end)


# 创建购物清单的字典，键为人的名字，值为购物清单的列表
shopping_lists = {}

for i in range(num_set_rows):
    # 分盒名称
    group_name = table_data.iloc[group_rows_start[i]-1,0]
    # 按行遍历
    for j in range(group_rows_start[i], group_rows_end[i]):
        row = table_data.iloc[j,:]
        item_name = row[0]
        item_pricead = row[1]
        item_quantity = row[2]
        # 只遍历范围内的cn
        for k in range(3,3+item_quantity):
            cn = row[k]
            if cn in shopping_lists:
                shopping_lists[cn]['数量'] += 1
                shopping_lists[cn]['调价'] += item_pricead
                # 分组
                if group_name in shopping_lists[cn]['明细']:
                    if item_name in shopping_lists[cn]['明细'][group_name]:
                        shopping_lists[cn]['明细'][group_name][item_name] +=1
                    else:
                        shopping_lists[cn]['明细'][group_name][item_name] =1
                else:
                    shopping_lists[cn]['明细'][group_name] = {item_name:1}
            else:
                shopping_lists[cn] = {'数量': 1, '调价': item_pricead, 
                                      '明细': {group_name: {item_name: 1}}}
# 明细字典改写为字符串
for cn in shopping_lists:
    item_quantity = shopping_lists[cn]['数量']
    item_pricead = shopping_lists[cn]['调价']
    for group_name in shopping_lists[cn]['明细']:
        item_str = ''
        for item_name in shopping_lists[cn]['明细'][group_name]:
            item_str += item_name + str(shopping_lists[cn]['明细'][group_name][item_name])
        shopping_lists[cn][group_name] = item_str
    shopping_lists[cn]['总价'] = price_average * item_quantity + item_pricead
    shopping_lists[cn]['cn'] = cn
    del shopping_lists[cn]['明细']
# print(shopping_lists)



# 将shopping_lists转换为DataFrame
shopping_lists_df = DataFrame(shopping_lists).T
# 按照指定顺序排列列
shopping_lists_df = shopping_lists_df[[
    'cn', '数量', '调价', 
    *[col for col in shopping_lists_df.columns if col not in ['cn', '数量', '调价', '总价']], 
    '总价','cn']]

print('清单生成完成')

# 将shopping_lists_df写入excel表格
writer = ExcelWriter('output.xlsx', engine='xlsxwriter')
shopping_lists_df.to_excel(writer, sheet_name='Sheet1', index=False)

# 获取工作簿和工作表对象
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# 设置表头格式
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'font_name': '等线',
    'align': 'center',
    'border': 1})

# 设置单数行格式
odd_format = workbook.add_format({
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#FFFFFF',
    'border': 1,
    'font_name': '等线'
    })

# 设置偶数行格式
even_format = workbook.add_format({
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#DDDDDD',
    'font_name': '等线',
    'border': 1})

# 设置表头格式
for col_num, value in enumerate(shopping_lists_df.columns.values):
    worksheet.write(0, col_num, value, header_format)

# 设置单数行和偶数行格式
for row_num in range(1, shopping_lists_df.shape[0]+1):
    if row_num % 2 == 0:
        format = even_format
    else:
        format = odd_format
    for col_num, value in enumerate(shopping_lists_df.iloc[row_num-1].values):
        if isna(value):
            worksheet.write(row_num, col_num, '', format)
        else:
            worksheet.write(row_num, col_num, value, format)

# 在最后一行的后面一行的第2、3、倒数第二列的单元格添加字符串
last_row = shopping_lists_df.shape[0]
last_col = shopping_lists_df.shape[1]
cell_sum = [1,2,last_col-2]
for i in cell_sum:
    cell_num = chr(65 +i)
    cell_num = '=sum(' + cell_num + '2:' + cell_num + str(last_row+1) + ')'
    worksheet.write(last_row +2 , i, cell_num, None)

copyright_str = 'Generated by GGTabulator beta version'
my_email = 'xinyi.bit@qq.com'
worksheet.write(last_row + 4 , 0, copyright_str, None)
worksheet.write(last_row + 5 , 0, 'E-mail:', None)
worksheet.write(last_row + 5 , 1, my_email, None)
# 关闭writer
writer.close()
print('成功输出结果到: \'output.xlsx\'')
print(copyright_str)
print('E-mail: ' + my_email)
input('按任意键退出程序')