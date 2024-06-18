# pyright:reportUnknownVariableType=false
# pyright:reportUnknownMemberType=false
# pyright:reportAttributeAccessIssue=false
# pyright:reportUnknownArgumentType=false
# pyright:reportMissingTypeStubs=false
import os
from collections import defaultdict

from pandas import DataFrame, read_excel


def file_not_found(filename: str) -> None:
    if filename == 'input.xlsx':
        print('未找到文件\'input.xlsx\'，请确保exe文件和表格文件处于同一目录下')
    else:
        print(f'未找到文件\'{filename}\'')
    key = input('按回车重启程序')
    if key == 'q':
        exit()
    else:
        return


def get_file_name(prompt:str) -> str:
    while True:
        filename = input(f'{prompt}\n>').strip(
            '"') or 'input.xlsx'
        if os.path.exists(filename):
            return filename
        else:
            file_not_found(filename)


def import_file(filename: str) -> DataFrame:
    table_data: DataFrame = read_excel(filename, header=None)

    def replace_symbol(x: str) -> str:
        return symbol_dict.get(x, x)

    # 校验table数据第一列
    invalid_rows = table_data[table_data.iloc[:, 0].isnull()].index
    if len(invalid_rows) > 0:
        print('\033[33m')
        print('警告: 输入文件中第{}行为无效行，删除后程序继续运行。'.format(list(invalid_rows+1)))
        table_data = table_data.drop(invalid_rows).reset_index(drop=True)
        print('      若结果有误，请再次检查输入文件并使第一列没有空的单元格。')
        print('\033[0m')

    # 试导入简称表
    try:
        symbol_data = read_excel(filename, header=None, sheet_name='symbol')
        # 将symbol表转换为字典，第1列为键，第2列为值
        symbol_dict = dict(zip(symbol_data.iloc[:, 0], symbol_data.iloc[:, 1]))
        # 在表中商品名称替换为简称，找不到对应的就不替换
        table_data.iloc[:, 0] = table_data.iloc[:, 0].apply(replace_symbol)
    except:
        print('未找到子表\'symbol\'用于替换简称，程序继续运行')
        return table_data

    return table_data


def import_paidfile(filename: str):
    paid_dict: dict[str, float] = defaultdict(float)
    try:
        paid_data = read_excel(filename, sheet_name='paid')
        # 将paid_data转换为字典，第1列为键，第2列为值。如果遇到重复的键，则将值相加
        for key, value in zip(paid_data.iloc[:, 0], paid_data.iloc[:, 1]):
            if key in paid_dict:
                paid_dict[key] += value
            else:
                paid_dict[key] = value
        # 判断paid表格是否为空
        if paid_dict == {}:
            print('识别到退补表但表格内容为空，不计算退补，程序继续运行')
        else:
            print('识别到退补表用于计算退补，程序继续运行')
    except:
        print('未找到退补表，程序继续运行')
    return paid_dict


def validate(lst: list[str]) -> list[str]:
    counts: dict[str, int] = defaultdict(int)
    validated_lst: list[str] = []
    for item in lst:
        counts[item] += 1
        if counts[item] > 1:
            validated_lst.append(f'{item}_{chr(48 + counts[item])}')
        else:
            validated_lst.append(item)
    return validated_lst


def get_column_letter(column_index: int) -> str:
    """
    将列号转换为Excel列的字母表示。
    例如，0 -> 'A', 1 -> 'B', ..., 26 -> 'AA'
    """
    letter = ''
    while column_index >= 0:
        letter = chr(column_index % 26 + 65) + letter
        column_index = column_index // 26 - 1
    return letter
