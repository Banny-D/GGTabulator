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
price_average_default = table_data.iloc[0,3]
print('默认均价：' + str(price_average_default))
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

# 导入退补表
paid_flag = 1
try:
    paid_data = read_excel('paid.xlsx')
    # 将paid_data转换为字典，第1列为键，第2列为值。如果遇到重复的键，则将值相加
    paid_dict = {}
    for key, value in zip(paid_data.iloc[:, 0], paid_data.iloc[:, 1]):
        if key in paid_dict:
            paid_dict[key] += value
        else:
            paid_dict[key] = value
    # 判断paid表格是否为空
    if paid_dict == {}:
        paid_flag = 0
        print('识别到文件\'paid.xlsx\'但表格内容为空，不计算退补，程序继续运行')
    else:
        print('识别到文件\'paid.xlsx\'用于计算退补，程序继续运行')
except FileNotFoundError:
    print('未找到文件\'paid.xlsx\'，程序继续运行')
    paid_flag = 0

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
# 创建均价表
price_average =[]

# 创建分盒顺序表，便于后续按照顺序输出分盒明细栏
group_order = []

print('校验：')

for i in range(num_set_rows):
    # 分盒名称
    group_name = table_data.iloc[group_rows_start[i]-1,0]
    # 更新分盒顺序表
    group_order.append(group_name)

    # 分盒均价
    group_price = table_data.iloc[group_rows_start[i]-1,3]
    if isna(group_price):
        price_average.append(price_average_default)
        group_price = price_average_default
    else:
        price_average.append(group_price)
    
    # 调价校验
    pricead = table_data.iloc[group_rows_start[i]:group_rows_end[i],1]
    quan = table_data.iloc[group_rows_start[i]:group_rows_end[i],2]
    sumpq = sum(pricead*quan)
    print('     '+group_name+':调价和为'+str(sumpq)+', 均价为'+str(price_average[i]))

    # 按行遍历
    for j in range(group_rows_start[i], group_rows_end[i]):
        row = table_data.iloc[j,:]
        item_name = row[0]
        item_pricead = row[1]
        item_quantity = row[2]
        # 商品数量校验
        if item_quantity<0:
            print('错误：第' + str(j+1) + '行商品数量错误')
            input('按任意键退出程序')
            exit()
        # 只遍历范围内的cn
        for k in range(3,3+item_quantity):
            cn = row[k]
            # cn数量少于商品数量的情况
            if isna(cn):
                print('错误：第' + str(j+1) + '行cn缺失')
                input('按任意键退出程序')
                exit()
            if cn in shopping_lists:
                
                if group_name in shopping_lists[cn]:
                    shopping_lists[cn][group_name]['数量'] += 1
                    shopping_lists[cn][group_name]['调价'] += item_pricead
                    if item_name in shopping_lists[cn][group_name]['明细']:
                        shopping_lists[cn][group_name]['明细'][item_name] +=1
                    else:
                        shopping_lists[cn][group_name]['明细'][item_name] =1
                else:
                    shopping_lists[cn][group_name] = {'数量': 1, '调价': item_pricead,
                                                    '明细': {item_name: 1}, '均价': group_price}
            else:
                shopping_lists[cn] = {group_name: {'数量': 1, '调价': item_pricead,
                                                   '明细': {item_name: 1}, '均价': group_price}}

print('统计完成')

# 计算退补
if paid_flag:
    for cn in paid_dict:
        if cn in shopping_lists:
            shopping_lists[cn]['已交'] = paid_dict[cn]
        else:
            shopping_lists[cn] = {'已交': paid_dict[cn]}

# 明细字典改写为字符串
# 并重构字典结构便于输出
for cn in shopping_lists:
    pricesum_str = '='
    quan_str = '='
    for group_name in shopping_lists[cn]:
        if group_name == '已交':
            continue
        item_quantity = shopping_lists[cn][group_name]['数量']
        item_pricead = shopping_lists[cn][group_name]['调价']
        item_str = ''
        group_price = shopping_lists[cn][group_name]['均价']
        del shopping_lists[cn][group_name]['数量']
        del shopping_lists[cn][group_name]['调价']
        del shopping_lists[cn][group_name]['均价']
        for item_name in shopping_lists[cn][group_name]['明细']:
            item_str += item_name + str(shopping_lists[cn][group_name]['明细'][item_name])
        del shopping_lists[cn][group_name]['明细']
        shopping_lists[cn][group_name] = item_str
        pricesum_str += '+' + str(group_price * item_quantity + item_pricead)
        quan_str += '+' + str(item_quantity)
    # 最后减去已交，用于计算退补
    if '已交' in shopping_lists[cn]:
        pricesum_str += '-' + str(shopping_lists[cn]['已交'])
    shopping_lists[cn]['总价'] = pricesum_str
    shopping_lists[cn]['总数'] = quan_str
    shopping_lists[cn]['cn'] = cn
    
# print(shopping_lists)



# 将shopping_lists转换为DataFrame
shopping_lists_df = DataFrame(shopping_lists).T
# 按照指定顺序排列列
if paid_flag:
    group_order.insert(0,'总数')
    group_order.insert(0,'cn')
    group_order.append('已交')
    group_order.append('总价')
    group_order.append('cn')
    shopping_lists_df = shopping_lists_df[group_order]
else:
    group_order.insert(0,'总数')
    group_order.insert(0,'cn')
    group_order.append('总价')
    group_order.append('cn')
    shopping_lists_df = shopping_lists_df[group_order]

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

# 添加条件格式
if paid_flag:
    highlight_format_positive = workbook.add_format({'font_color': '#FFA500'})  # Orange color for positive values
    highlight_format_negative = workbook.add_format({'font_color': '#4F81BD'})  # Blue color for zero or negative values
    # 获取倒数第二列的列号
    second_last_col = shopping_lists_df.shape[1] - 2
    worksheet.conditional_format('{}2:{}{}'.format(chr(65 + second_last_col),
                                                    chr(65 + second_last_col),
                                                    shopping_lists_df.shape[0]+1), 
                                    {'type': 'cell',
                                    'criteria': '>',
                                    'value': 0,
                                    'format': highlight_format_positive})
    worksheet.conditional_format('{}2:{}{}'.format(chr(65 + second_last_col),
                                                    chr(65 + second_last_col),
                                                    shopping_lists_df.shape[0]+1), 
                                    {'type': 'cell',
                                    'criteria': '<=',
                                    'value': 0,
                                    'format': highlight_format_negative})

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