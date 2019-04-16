# -*- coding: utf-8 -*-
import pandas as pd
import pathlib
from typing import List, Dict


def format_progress(msg: str, no_return: bool = False) -> None:
    if no_return:
        print(msg, end='')
    else:
        print(msg)


def charge_off_amount(amount: pd.Series, sign: pd.Series,
                      charge_off_words: List[str]) -> None:
    amount[sign.isin(charge_off_words)] *= -1


def get_account_name(name: str, deco_strings: str or List[str]) -> str:
    if isinstance(deco_strings, str):
        name = name.replace(deco_strings, '')
    elif isinstance(deco_strings, List[str]):
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


class BankPara:
    def __init__(self,
                 col_order: List[str],
                 col_names: List[str],
                 col_rename: Dict[str, str] = None,
                 deco_strings: str or List[str] = None):
        # 列顺序，使用原文件中列名
        self.col_order = col_order
        # 最终列名称，从下面选择，不要改变顺序且需与列顺序相同
        # words = [
        #     '银行名称', '户名', '账号', '交易日期', '交易方式', '收付标志', '来源和用途', '币种',
        #     '金额(原币)', '金额(美元)', '余额', '对手户名', '对手账号', '对手开户行', '交易地区', '交易场所',
        #     '交易网点', '柜员号', '涉外交易代码', '交易代码', '代办人', '代办人证件', '备注', '附言', '其他'
        # ]
        self.col_names = col_names
        # 需要修改的列名字典{'原列名':'新列名'}，主要解决同一银行不同文件列名不同的问题。
        self.col_rename = col_rename
        # 在包含户名的文件名中将要被截取掉的多余字串，文件中包含户名传入None，否则需要传入需从文件名截取掉的字串，不需截取转入空串
        self.deco_strings = deco_strings


def parse_transaction(dir_path: pathlib.Path, test_header: int,
                      charge_off_words: List[str],
                      bank_paras: BankPara) -> pd.DataFrame:
    format_progress('开始分析' + dir_path.name + '账户……')
    tmp_trans_list_by_name = []
    tmp_all_nums = 0
    for trans_file in dir_path.iterdir():  # 对每一个文件
        if trans_file.match('~*') or trans_file.match('*账户情况.xls*'):
            continue
        format_progress('    ' + trans_file.name + '……', True)
        excel_file = pd.ExcelFile(trans_file)
        tmp_trans_list_by_account = []
        tmp_not_parsed = []
        tmp_line_num = 0
        for sheet in excel_file.sheet_names:  # 对每一个工作表
            header = get_header(excel_file, sheet, test_header)
            if header == -1:  # 空工作表
                continue
            elif header == -2:  # 含数据但无法解析工作表
                tmp_not_parsed.append(sheet)
                continue
            else:
                tmp_trans_sheet = excel_file.parse(sheet_name=sheet,
                                                   header=header,
                                                   dtype=str)
                if bank_paras.col_rename is not None:
                    tmp_trans_sheet.rename(columns=bank_paras.col_rename,
                                           inplace=True)
                tmp_trans_list_by_account.append(tmp_trans_sheet)
                tmp_line_num += len(tmp_trans_sheet)
                continue
        tmp_transactions = pd.concat(tmp_trans_list_by_account, sort=False)
        if bank_paras.deco_strings is not None:
            tmp_transactions['户名'] = get_account_name(trans_file.stem,
                                                      bank_paras.deco_strings)
        tmp_trans_list_by_name.append(tmp_transactions)
        format_progress('工作表解析成功' + str(len(tmp_trans_list_by_account)) +
                        '/失败' + str(len(tmp_not_parsed)) + '，解析流水' +
                        str(tmp_line_num) + '条')
        tmp_all_nums += tmp_line_num
    transactions = pd.concat(tmp_trans_list_by_name, sort=False)
    transactions['银行名称'] = dir_path.name
    transactions = transactions.reindex(columns=bank_paras.col_order)
    transactions.columns = bank_paras.col_names
    transactions.dropna(axis=0,
                        how='any',
                        subset=['交易日期', '金额(原币)'],
                        inplace=True)
    transactions['交易日期'] = pd.to_datetime(transactions['交易日期'])
    transactions['金额(原币)'] = pd.to_numeric(transactions['金额(原币)'])
    charge_off_amount(transactions['金额(原币)'], transactions['收付标志'],
                      charge_off_words)
    format_progress('    分析结束，共解析' + str(len(transactions)) + '/' +
                    str(tmp_all_nums) + '条')
    return transactions
