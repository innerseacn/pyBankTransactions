# -*- coding: utf-8 -*-
# %%
import pandas as pd
import pathlib

base_path = pathlib.Path(r'C:\Users\InnerSea\工作\反洗钱\附件：罗凯声等人账户交易明细')


def format_progress(msg: str):
    print(msg)


# %%
def parse_accounts(dir_path: pathlib.Path):
    acc_file = next(dir_path.glob(dir_path.name + '账户情况.*'))
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


def format_account_list(base_path: pathlib.Path):
    format_progress('开始分析账户列表……')
    tmp_account_list = []
    for dir in base_path.iterdir():
        if dir.is_dir():
            tmp_account_list.append(parse_accounts(dir))
    account_list = pd.concat(tmp_account_list, ignore_index=True)
    account_list = account_list[[
        '银行', '户名', '证件号码', '开户留存基本信息', '账号', '开户行', '开户日期', '当前余额', '账户性质',
        '账户状态', '备注'
    ]]
    return account_list


def format_boc_account_map(dir_path: pathlib.Path):
    format_progress('开始生成中国银行账号对应表……')
    map_file = next(dir_path.glob('*交易*.*'))
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
    account_map = pd.concat(map_list, ignore_index=False)
    account_map = account_map[[
        '姓名', '证件', '卡号', '账号', '对应新账号', '主账号', '子账号', '货币'
    ]]
    account_map.dropna(axis=0, how='any', subset=['姓名'], inplace=True)
    return account_map


def format_bjyh(dir_path: pathlib.Path):
    format_progress('开始分析北京银行账户……')
    trans_file = next(dir_path.glob('*交易*.*'))
    format_progress(trans_file.name)
    tmp_trans_list = pd.read_excel(
        trans_file,
        sheet_name=None,
        converters={
            '帐号': str,
            '交易日期': lambda x: pd.to_datetime(str(x)),
            '交易对手帐号': str,
        })
    for name in tmp_trans_list:
        tmp_trans_list[name]['户名'] = name
    transactions = pd.concat(tmp_trans_list.values())
    transactions['银行名称'] = "北京银行"
    tmp_t = transactions['资金收付标志']
    transactions['金额'][(tmp_t == '付') | (tmp_t == '支出')] *= -1
    transactions = transactions[[
        '银行名称', '户名', '帐号', '交易日期', '交易方式', '资金收付标志', '金额', '余额', '交易附言',
        '交易对手姓名', '交易对手帐号', '交易对手金融机构名称'
    ]]

    return transactions


# %%
def format_bhyh(dir_path: pathlib.Path):
    format_progress('开始分析渤海银行账户……')

    if dir_path.is_file():
        raise Exception('渤海银行账户应该是一个文件夹！')

    accounts = parse_accounts(dir_path)

    tmp_trans_list = []
    for trans_file in dir_path.glob('*交易*明细*.*'):
        format_progress(trans_file.name)
        tmp_transactions = pd.read_excel(
            trans_file,
            header=1,
            converters={
                '账号': str,
                '交易日期': lambda x: pd.to_datetime(str(x)),
                '交易对手账号': str,
            })
        tmp_transactions['户名'] = trans_file.stem
        tmp_trans_list.append(tmp_transactions)
    transactions = pd.concat(tmp_trans_list)

    accounts['银行名称'] = "渤海银行"
    transactions['银行名称'] = "渤海银行"
    tmp_t = transactions['资金收付标志']
    transactions[' 交易额(按原币计)'][(tmp_t == '借') | (tmp_t == '借方')] *= -1
    transactions[' 交易额(折合美元)'][(tmp_t == '借') | (tmp_t == '借方')] *= -1
    transactions = transactions[[
        '户名', '银行名称', '金融机构名称', '账号', '交易日期', '交易方式', '涉外收支交易分类与代码', '资金收付标志',
        '币种', ' 交易额(按原币计)', ' 交易额(折合美元)', '交易对手姓名或名称', '交易对手账号', '业务标示号',
        '代办人姓名', '代办人身份证件/证明文件号码', '备注'
    ]]

    return accounts, transactions


accounts, transactions = format_bhyh(base_path / '渤海银行')

# %%
accounts
# %%
transactions

# %%
account_list.append(accounts)
transaction_list.append(transactions)
