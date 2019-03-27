# -*- coding: utf-8 -*-
# %%
import pandas as pd
import pathlib
from typing import List

base_path = pathlib.Path(r'C:\Users\InnerSea\工作\反洗钱\附件：罗凯声等人账户交易明细')


def format_progress(msg: str):
    print(msg)


def get_account_name(name: str, deco_strings: str or List[str] = None) -> str:
    if isinstance(deco_strings, str):
        name = name.replace(deco_strings, '')
    elif isinstance(deco_strings, List[str]):
        for string in deco_strings:
            name = name.replace(string, '')
    return name


def parse_accounts(dir_path: pathlib.Path) -> pd.DataFrame:
    acc_file = next(dir_path.glob(dir_path.name + '账户情况.xls*'))
    format_progress(acc_file.name)
    accounts = pd.read_excel(
        acc_file, header=1, dtype={
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
    account_list = account_list[[
        '银行', '户名', '证件号码', '开户留存基本信息', '账号', '开户行', '开户日期', '当前余额', '账户性质',
        '账户状态', '备注'
    ]]
    return account_list


def format_boc_account_map(dir_path: pathlib.Path) -> pd.DataFrame:
    format_progress('开始生成中国银行账号对应表……')
    map_file = next(dir_path.glob('*交易*.xls*'))
    map_list = pd.read_excel(
        map_file,
        sheet_name=['旧线账号', '新线账号'],
        dtype={
            '证件': str,
            '账号': str,
            '卡号': str,
            '对应新账号': str,
            '主账号': str,
            '子账号': str
        })
    account_map = pd.concat(map_list, ignore_index=False, sort=False)
    account_map = account_map[[
        '姓名', '证件', '卡号', '账号', '对应新账号', '主账号', '子账号', '货币'
    ]]
    account_map.dropna(axis=0, how='any', subset=['姓名'], inplace=True)
    return account_map


def format_bjyh(dir_path: pathlib.Path) -> pd.DataFrame:
    format_progress('开始分析北京银行账户……')
    trans_file = next(dir_path.glob('*交易*.xls*'))
    format_progress(trans_file.name)
    tmp_trans_list = pd.read_excel(
        trans_file,
        sheet_name=None,
        dtype={
            '帐号': str,
            '交易日期': str,
            '交易对手帐号': str
        })
    for name in tmp_trans_list:
        tmp_trans_list[name]['户名'] = name
    transactions = pd.concat(tmp_trans_list.values(), sort=False)
    transactions['银行名称'] = "北京银行"
    tmp_t = transactions['资金收付标志']
    transactions['金额'][(tmp_t == '付') | (tmp_t == '支出')] *= -1
    transactions = transactions[[
        '银行名称', '户名', '帐号', '交易日期', '交易方式', '资金收付标志', '金额', '余额', '交易附言',
        '交易对手姓名', '交易对手帐号', '交易对手金融机构名称'
    ]]
    transactions.columns = [
        '银行名称', '户名', '账号', '交易日期', '交易方式', '资金收付标志', ' 交易额(按原币计)', '账户余额',
        '交易附言', '交易对手姓名或名称', '交易对手账号', '交易对手金融机构名称'
    ]
    return transactions


def format_common(dir_path: pathlib.Path,
                  deco_strings: str or List[str] = None) -> pd.DataFrame:
    format_progress('开始分析' + dir_path.name + '账户……')
    tmp_trans_list_by_name = []
    for trans_file in dir_path.iterdir():
        if trans_file.match('~*') or trans_file.match('*账户情况.xls*'):
            continue
        format_progress(trans_file.name)
        tmp_trans_list_by_account = pd.read_excel(
            trans_file,
            header=1,
            sheet_name=None,
            dtype={
                '账号': str,
                '交易日期': str,
                '交易对手账号': str
            })
        tmp_transactions = pd.concat(tmp_trans_list_by_account, sort=False)
        tmp_transactions['户名'] = get_account_name(trans_file.stem,
                                                  deco_strings)
        tmp_trans_list_by_name.append(tmp_transactions)
    transactions = pd.concat(tmp_trans_list_by_name, sort=False)
    transactions['银行名称'] = dir_path.name
    tmp_t = transactions['资金收付标志']
    transactions[' 交易额(按原币计)'][(tmp_t == '借') | (tmp_t == '借方')] *= -1
    transactions[' 交易额(折合美元)'][(tmp_t == '借') | (tmp_t == '借方')] *= -1
    transactions = transactions[[
        '银行名称', '户名', '账号', '交易日期', '交易方式', '资金收付标志', '币种', ' 交易额(按原币计)',
        ' 交易额(折合美元)', '交易对手姓名或名称', '交易对手账号', '涉外收支交易分类与代码', '业务标示号', '代办人姓名',
        '代办人身份证件/证明文件号码', '备注'
    ]]
    return transactions


def format_transactions(base_path: pathlib.Path) -> pd.DataFrame:
    format_progress('开始分析银行流水……')
    tmp_trans_list_by_bank = []
    for dir in base_path.iterdir():
        if dir.is_dir():
            if dir.name == '北京银行':
                tmp_trans_list_by_bank.append(format_bjyh(dir))
            elif dir.name == '渤海银行':
                tmp_trans_list_by_bank.append(
                    format_common(dir, '报告可疑交易逐笔明细表—'))
            else:
                pass
    transactions = pd.concat(
        tmp_trans_list_by_bank, ignore_index=True, sort=False)
    transactions.dropna(axis=0, how='any', subset=['交易日期'], inplace=True)
    return transactions
