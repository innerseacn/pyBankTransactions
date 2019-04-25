# -*- coding: utf-8 -*-
import pandas as pd
import pathlib
import statics as st
from typing import List, Union


def format_progress(msg: str, no_return: bool = False) -> None:
    if no_return:
        print(msg, end='')
    else:
        print(msg)


# 将出账金额置为复数，方便观看。但要注意有些冲抵金额本身为负，
# 应用本函数后变为正数，所以自动化计算需要结合收付标志处理。
def charge_off_amount(amount: pd.Series, sign: pd.Series) -> None:
    amount[sign.isin(st.CHARGE_OFF_WORDS)] *= -1


def get_account_name(name: str, deco_strings: Union[str, List[str]]) -> str:
    if isinstance(deco_strings, str):
        name = name.replace(deco_strings, '')
    elif isinstance(deco_strings, list):
        for string in deco_strings:
            name = name.replace(string, '')
    return name


def get_header(excel_file: pd.ExcelFile, sheet: str, test_header: int) -> int:
    for header in range(0, test_header):  # 尝试解析标题行TEST_HEADER次
        tmp_trans_sheet = excel_file.parse(sheet_name=sheet,
                                           nrows=0,
                                           header=header)
        if len(tmp_trans_sheet.columns) == 0:
            return -1
        if 'Unnamed: 1' not in tmp_trans_sheet.columns:  # 一旦找到标题行则返回行号
            return header
    return -2


# 在入账出账单独成列时，将第二列合并到第一列
def combine_amount_cols(data: pd.DataFrame, second_amount_col: str) -> None:
    data['交易金额'][pd.to_numeric(data['交易金额']) == 0] = data[second_amount_col]


def parse_transaction(dir_path: pathlib.Path,
                      bank_para: st.BankPara) -> pd.DataFrame:
    format_progress('开始分析{}账户……'.format(dir_path.name))
    tmp_trans_list_by_name = []  # 流水列表（按文件）
    tmp_all_nums = 0  # 所有流水行数
    for trans_file in dir_path.iterdir():  # 对每一个文件
        if trans_file.match('~*') or trans_file.match('*账户情况.xls*'):
            continue  # 如是无用文件则跳过
        format_progress('    {}……'.format(trans_file.name), True)
        excel_file = pd.ExcelFile(trans_file)
        tmp_trans_list_by_account = []  # 当前文件流水列表（按工作表）
        tmp_not_parsed = 0  # 当前文件无法解析工作表数
        tmp_line_num = 0  # 当前文件流水行数
        for sheet in excel_file.sheet_names:  # 对每一个工作表
            header = get_header(excel_file, sheet, st.TEST_HEADER)  # 寻找表头
            if header == -1:  # 空工作表
                continue
            elif header == -2:  # 含数据但无法解析的工作表
                tmp_not_parsed += 1
                continue
            else:  # 找到表头
                tmp_trans_sheet = excel_file.parse(sheet_name=sheet,
                                                   header=header,
                                                   dtype=str)
                tmp_trans_sheet.rename(columns=bank_para.col_map, inplace=True)
                # 如果本文件名符合如下规则, 此时认为工作表名就是户名，若deco_strings为空则增加户名列
                if trans_file.match(
                        '交易明细模板*.xls*') and bank_para.sheet_name_is_acc_name:
                    tmp_trans_sheet['户名'] = sheet
                tmp_trans_list_by_account.append(tmp_trans_sheet)
                tmp_line_num += len(tmp_trans_sheet)
                continue
        tmp_transactions = pd.concat(tmp_trans_list_by_account, sort=False)
        try:
            tmp_transactions.dropna(axis=0,
                                    how='any',
                                    subset=['交易日期', '交易金额'],
                                    inplace=True)
            if bank_para.deco_strings is not None:  # 如果流水中不包含户名列，则此项不为空
                tmp_transactions['户名'] = get_account_name(trans_file.stem,
                                                        bank_para.deco_strings)
            tmp_trans_list_by_name.append(tmp_transactions)
        except KeyError:
            format_progress('工作表映射错误，跳过文件\n                ', True)
        format_progress('工作表解析成功{}/失败{}，解析流水{}/{}条'.format(
            len(tmp_trans_list_by_account), tmp_not_parsed,
            len(tmp_transactions), tmp_line_num))
        tmp_all_nums += tmp_line_num
    transactions = pd.concat(tmp_trans_list_by_name, sort=False)
    transactions['银行名称'] = dir_path.name
    if bank_para.second_amount_col is not None:
        combine_amount_cols(transactions, bank_para.second_amount_col)
    # transactions = transactions.reindex(columns=st.COLUMN_ORDER)
    transactions['借贷标志'] = transactions['借贷标志'].map(str.strip)
    transactions['交易日期'] = pd.to_datetime(transactions['交易日期'])
    transactions['交易金额'] = pd.to_numeric(transactions['交易金额'])
    if not bank_para.has_minus_amounts:
        charge_off_amount(transactions['交易金额'], transactions['借贷标志'])
    format_progress('    分析结束，共解析{}/{}条'.format(len(transactions),
                                                tmp_all_nums))
    if len(transactions) < tmp_all_nums:
        format_progress(
            '                   ^----------------------------- 请查找问题，若无问题请忽略')
    return transactions
