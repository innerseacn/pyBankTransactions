# -*- coding: utf-8 -*-
import pandas as pd
import core
import pathlib
import statics as st
from typing import List, Union


# 以下为账户分析函数
def parse_accounts(dir_path: pathlib.Path) -> pd.DataFrame:
    acc_file = next(dir_path.glob(dir_path.name + '账户情况.xls*'))
    format_progress(acc_file.name)
    accounts = pd.read_excel(acc_file,
                             header=1,
                             dtype={
                                 '证件号码': str,
                                 '账号': str,
                                 '开户日期': str
                             })
    accounts.dropna(axis=0, how='any', subset=['账号'], inplace=True)
    accounts.dropna(axis=1, how='all', inplace=True)
    accounts['户名'].fillna(method='ffill', inplace=True)
    accounts['银行'] = dir_path.name
    return accounts


def format_account_list(base_path: pathlib.Path) -> pd.DataFrame:
    format_progress('开始分析账户列表……')
    tmp_account_list = []
    for dir in base_path.iterdir():
        if dir.is_dir():
            tmp_account_list.append(parse_accounts(dir))
    account_list = pd.concat(tmp_account_list, ignore_index=True, sort=False)
    account_list = account_list.reindex(columns=[
        '银行', '户名', '证件号码', '开户留存基本信息', '账号', '开户行', '开户日期', '当前余额', '账户性质',
        '账户状态', '备注'
    ])
    return account_list


def format_boc_account_map(dir_path: pathlib.Path) -> pd.DataFrame:
    format_progress('开始生成中国银行账号对应表……')
    map_file = next(dir_path.glob('*交易*.xls*'))
    excel_file = pd.ExcelFile(map_file)
    map_list = excel_file.parse(sheet_name=['旧线账号', '新线账号'],
                                dtype={
                                    '证件': str,
                                    '账号': str,
                                    '卡号': str,
                                    '对应新账号': str,
                                    '主账号': str,
                                    '子账号': str
                                })
    account_map = pd.concat(map_list, ignore_index=False, sort=False)
    account_map = account_map.reindex(
        columns=['姓名', '证件', '卡号', '账号', '对应新账号', '主账号', '子账号', '货币'])
    account_map.dropna(axis=0, how='any', subset=['姓名'], inplace=True)
    return account_map


# 以下为流水分析函数
def format_jsyh(dir_path: pathlib.Path) -> pd.DataFrame:
    pass  # 建设银行太复杂回头在写


def format_transactions(base_path: pathlib.Path) -> pd.DataFrame:
    core.format_progress('开始分析银行流水……')
    tmp_trans_list_by_bank = []
    tmp_banks_no_support = 0
    for dir in base_path.iterdir():
        try:
            if dir.is_dir():
                tmp_trans_list_by_bank.append(
                    core.parse_transaction(dir, st.BANK_PARAS[dir.name]))
        except KeyError as k:
            tmp_banks_no_support += 1
            core.format_progress('暂不支持{}'.format(k))
    transactions = pd.concat(tmp_trans_list_by_bank,
                             ignore_index=True,
                             sort=False)
    transactions = transactions.reindex(columns=st.COLUMN_ORDER)
    core.format_progress('全部分析完成，\n    成功解析银行{}家，流水{}条\n    发现暂不支持银行{}家'.format(
        len(tmp_trans_list_by_bank), len(transactions), tmp_banks_no_support))
    return transactions
