from ggt_functions import *
from pandas import DataFrame, ExcelWriter, isna, read_excel

copyright_str = 'Generated by GGTabulator beta version'
my_email = 'xinyi.bit@qq.com'


def main():

    file_name = get_file_name()
    # 导入表格
    table_data = import_file(file_name)

    # 导入退补表
    paid_dict = import_paidfile(file_name)

    # 均价
    price_average_default = table_data.iloc[0, 3]
    if isna(price_average_default):
        price_average_default = 0
    print('默认均价：{:.2f}'.format(price_average_default))

    # 找出分盒的开始位置和结束位置
    group_rows = [0] + table_data[table_data.iloc[:, 1].isnull()
                                  ].index.tolist()
    group_start = [i+1 for i in group_rows]
    group_end = [i for i in group_rows[1:]] + [table_data.shape[0]]
    # 分盒的数量
    num_set_rows = len(group_rows)

    # 创建购物清单的字典，键为人的名字，值为购物清单的列表
    shopping_lists = {}
    # 创建均价表
    price_average = []

    # 创建分盒顺序表，便于后续按照顺序输出分盒明细栏
    group_name_order = table_data.iloc[group_rows, 0].tolist()
    group_order_full = []
    # 校验并修正使分盒名称不重复
    group_name_order = validate(group_name_order)

    print('校验：')

    for i in range(num_set_rows):
        # 分盒名称
        group_name = group_name_order[i]
        # 更新分盒顺序表
        group_order_full.append(group_name)
        group_order_full.append(group_name+'总数')
        group_order_full.append(group_name+'总价')

        # 分盒均价
        group_price = table_data.iloc[group_rows[i], 3]
        if isna(group_price):
            group_price = price_average_default
        price_average.append(group_price)

        # 按行遍历
        for j in range(group_start[i], group_end[i]):
            row = table_data.iloc[j, :]
            item_name = str(row[0])
            item_pricead = row[1]
            item_quantity = row[2]

            # 商品名称校验
            if item_name[0].isdigit() or item_name[-1].isdigit():
                # 在以数字开头或结尾的商品名称前后增加标识符
                item_name = '/' + item_name + ':'

            # 商品数量校验
            if item_quantity < 0:
                print('错误：第' + str(j+1) + '行商品数量错误')
                print('请检查调价列与数量列的顺序是否有误')
                return 0
            elif not isinstance(item_quantity, int):
                print('错误：第' + str(j+1) + '行商品数量错误')
                print('请检查该行商品数量是否为空或不是整数')
                return 0
            # 只遍历范围内的cn
            for k in range(3, 3+item_quantity):
                cn = row[k]
                # cn数量少于商品数量的情况
                if isna(cn):
                    print('错误：第' + str(j+1) + '行cn缺失')
                    print('请检查cn列表, 配比数范围内不要有空的单元格')
                    return 0
                # cn存在于shoppinglist时
                if cn in shopping_lists:
                    if group_name in shopping_lists[cn]:
                        shopping_lists[cn][group_name + '总数'] += 1
                        shopping_lists[cn][group_name +
                                           '总价'] += group_price + item_pricead
                        if item_name in shopping_lists[cn][group_name]:
                            shopping_lists[cn][group_name][item_name] += 1
                        else:
                            shopping_lists[cn][group_name][item_name] = 1
                    else:
                        shopping_lists[cn][group_name] = {item_name: 1}
                        shopping_lists[cn][group_name + '总数'] = 1
                        shopping_lists[cn][group_name +
                                           '总价'] = group_price + item_pricead
                # cn没有被统计过，新建一个以cn为键的键值对
                else:
                    # shopping_lists[cn] = {group_name: {'数量': 1, '调价': item_pricead,
                    #                                    '明细': {item_name: 1}, '均价': group_price}}
                    shopping_lists[cn] = {group_name: {item_name: 1}, group_name+'总数': 1,
                                          group_name+'总价': group_price + item_pricead}
        # 调价校验
        pricead = table_data.iloc[group_start[i]:group_end[i], 1]
        quan = table_data.iloc[group_start[i]:group_end[i], 2]
        sumpq = sum(pricead*quan)
        print('     '+group_name+'：调价和为{:.2f}，均价为{:.2f}，分盒总价为{:.2f}'
              .format(sumpq, price_average[i], sum((group_price + pricead)*quan)))

    print('统计完成')

    # 计算退补
    if paid_dict != {}:
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
        for group_name in group_name_order:
            if not (group_name in shopping_lists[cn]):
                continue

            item_quantity = shopping_lists[cn][group_name + '总数']
            item_pricead = shopping_lists[cn][group_name + '总价']
            item_str = ''

            for item_name in shopping_lists[cn][group_name]:
                item_str += item_name + \
                    str(shopping_lists[cn][group_name][item_name])
            shopping_lists[cn][group_name] = item_str
            pricesum_str += '+' + '{:.2f}'.format(item_pricead)
            quan_str += '+' + '{:.2f}'.format(item_quantity)

        shopping_lists[cn]['总价'] = pricesum_str
        shopping_lists[cn]['总数'] = quan_str
        shopping_lists[cn]['cn'] = cn
        shopping_lists[cn]['蓝退红补'] = pricesum_str
        # 最后减去已交，用于计算退补
        if '已交' in shopping_lists[cn]:
            pricesum_str += '-' + '{:.2f}'.format(shopping_lists[cn]['已交'])
            shopping_lists[cn]['蓝退红补'] = pricesum_str

    # 将shopping_lists转换为DataFrame
    shopping_lists_df = DataFrame(shopping_lists).T
    # 按照指定顺序排列列
    group_order_full.insert(0, 'cn')
    if num_set_rows > 1:
        group_order_full.append('总数')
        group_order_full.append('总价')
    if paid_dict != {}:
        group_order_full.append('已交')
        group_order_full.append('蓝退红补')
    group_order_full.append('cn')
    shopping_lists_df = shopping_lists_df[group_order_full]

    print('清单生成完成')

    # 将shopping_lists_df写入excel表格

    while True:
        try:
            writer = ExcelWriter('output.xlsx', engine='xlsxwriter')
            break
        except:
            input('\033[31m错误：无法读取和写入，请先关闭或删除\'output.xlsx\'再继续运行程序。\033[0m')

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
    worksheet.freeze_panes(1, 0)  # 冻结第一行

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
    if paid_dict != {}:
        highlight_format_positive = workbook.add_format(
            {'bg_color': '#FFC7CE', 'font_color': '#9C0006'})  # Orange color for positive values
        highlight_format_negative = workbook.add_format(
            {'bg_color': '#DCE6F1', 'font_color': '#00008B'})  # Blue color for zero or negative values
        # 获取倒数第二列的列号
        second_last_col = shopping_lists_df.shape[1] - 2
        worksheet.conditional_format('{}2:{}{}'.format(get_column_letter(second_last_col),
                                                       get_column_letter(
                                                           second_last_col),
                                                       shopping_lists_df.shape[0]+1),
                                     {'type': 'cell',
                                         'criteria': '>',
                                         'value': 0,
                                         'format': highlight_format_positive})
        worksheet.conditional_format('{}2:{}{}'.format(get_column_letter(second_last_col),
                                                       get_column_letter(
                                                           second_last_col),
                                                       shopping_lists_df.shape[0]+1),
                                     {'type': 'cell',
                                         'criteria': '<=',
                                         'value': 0,
                                         'format': highlight_format_negative})

    # 在最后一行的后面一行的第2、3、倒数第二列的单元格添加字符串
    last_row = shopping_lists_df.shape[0]
    last_col = shopping_lists_df.shape[1]
    cell_sum = range(2, last_col-1)
    for i in cell_sum:
        cell_num = get_column_letter(i)
        cell_num = '=sum(' + cell_num + '2:' + cell_num + str(last_row+1) + ')'
        worksheet.write(last_row + 2, i, cell_num, None)

    worksheet.write(last_row + 4, 0, copyright_str, None)
    worksheet.write(last_row + 5, 0, 'E-mail:', None)
    worksheet.write(last_row + 5, 1, my_email, None)
    # 关闭writer
    writer.close()
    print('成功输出结果到: \'output.xlsx\'')
    print(copyright_str)
    print('E-mail: ' + my_email)
    return 1


while True:
    main()
    in_key = input('按回车重启程序')
    if in_key == 'q':
        break
