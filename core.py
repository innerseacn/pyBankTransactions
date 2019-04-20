# -*- coding: utf-8 -*-
import pandas as pd
import pathlib
import statics as st
from typing import List, Dict, Union


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
def combine_amount_cols(data: pd.DataFrame, first_col: str,
                        second_col: str) -> None:
    data[first_col][pd.to_numeric(data[first_col]) == 0] = data[second_col]


class BankPara:
    def __init__(self,
                 col_order: List[str],
                 col_names: List[str],
                 col_rename: Dict[str, str] = None,
                 has_two_amount_cols: List[str] = None,
                 deco_strings: Union[str, List[str]] = None) -> None:
        # 列顺序，使用原文件中列名，顺序依照dataFormatting.COLUMN_NAMES
        self.col_order = col_order
        # 最终列名称，从dataFormatting.COLUMN_NAMES中选择，不要改变顺序
        self.col_names = col_names
        # 需要修改的列名字典{'原列名':'新列名'}，主要解决同一银行不同文件列名不同的问题。
        self.col_rename = col_rename
        # 当流水中入账和出账各自成一列时，此项赋两列的列名，否则为None
        self.has_two_amount_cols = has_two_amount_cols
        # 在包含户名的文件名中将要被截取掉的多余字串，文件中包含户名传入None，否则需要传入需从文件名截取掉的字串，不需截取转入空串
        self.deco_strings = deco_strings


def parse_transaction(dir_path: pathlib.Path, test_header: int,
                      bank_paras: BankPara) -> pd.DataFrame:
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
            header = get_header(excel_file, sheet, test_header)  # 寻找表头
            if header == -1:  # 空工作表
                continue
            elif header == -2:  # 含数据但无法解析的工作表
                tmp_not_parsed += 1
                continue
            else:  # 找到表头
                tmp_trans_sheet = excel_file.parse(sheet_name=sheet,
                                                   header=header,
                                                   dtype=str)
                if bank_paras.col_rename is not None:
                    tmp_trans_sheet.rename(columns=bank_paras.col_rename,
                                           inplace=True)
                # 如果本文件名符合如下规则, 此时认为工作表名就是户名，若deco_strings为空则增加户名列
                if trans_file.match('*交易*明细*.xls*') and (
                        bank_paras.deco_strings is None):
                    tmp_trans_sheet['户名'] = sheet
                tmp_trans_list_by_account.append(tmp_trans_sheet)
                tmp_line_num += len(tmp_trans_sheet)
                continue
        tmp_transactions = pd.concat(tmp_trans_list_by_account, sort=False)
        if bank_paras.deco_strings is not None:  # 如果流水中不包含户名列，则此项不为空
            tmp_transactions['户名'] = get_account_name(trans_file.stem,
                                                      bank_paras.deco_strings)
        tmp_trans_list_by_name.append(tmp_transactions)
        format_progress('工作表解析成功{}/失败{}，解析流水{}条'.format(
            len(tmp_trans_list_by_account), tmp_not_parsed, tmp_line_num))
        tmp_all_nums += tmp_line_num
    transactions = pd.concat(tmp_trans_list_by_name, sort=False)
    transactions['银行名称'] = dir_path.name
    if bank_paras.has_two_amount_cols is not None:
        combine_amount_cols(transactions, bank_paras.has_two_amount_cols[0],
                            bank_paras.has_two_amount_cols[1])
    transactions = transactions.reindex(columns=bank_paras.col_order)
    transactions.columns = bank_paras.col_names
    transactions.dropna(axis=0,
                        how='any',
                        subset=['交易日期', '金额(原币)'],
                        inplace=True)
    transactions['收付标志'] = transactions['收付标志'].map(str.strip)
    transactions['交易日期'] = pd.to_datetime(transactions['交易日期'])
    transactions['金额(原币)'] = pd.to_numeric(transactions['金额(原币)'])
    charge_off_amount(transactions['金额(原币)'], transactions['收付标志'])
    format_progress('    分析结束，共解析{}/{}条'.format(len(transactions),
                                                tmp_all_nums))
    if len(transactions) < tmp_all_nums:
        format_progress(
            '                   ^----------------------------- 请查找问题')
    return transactions
