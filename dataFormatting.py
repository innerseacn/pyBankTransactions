# -*- coding: utf-8 -*-
# %%
import pandas as pd
import pathlib
base_path = pathlib.Path(r'C:\Users\InnerSea\工作\反洗钱\附件：罗凯声等人账户交易明细')

# %%
account_list = []
transaction_list = []


# %%
def format_progress(msg: str):
    print(msg)


# %%
def format_bjyh(dir_path: pathlib.Path):
    format_progress('开始分析北京银行账户……')

    if dir_path.is_file():
        raise Exception('北京银行账户应该是一个文件夹！')

    acc_file = next(dir_path.glob('*账户*.*'))
    format_progress(acc_file.name)
    accounts = pd.read_excel(
        acc_file,
        header=1,
        converters={'开户日期': lambda x: pd.to_datetime(str(x))})

    trans_file = next(dir_path.glob('*交易明细*.*'))
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

    accounts['银行名称'] = "北京银行"
    transactions['银行名称'] = "北京银行"
    tmp_t = transactions['资金收付标志']
    transactions['金额'][(tmp_t == '付') | (tmp_t == '支出')] *= -1
    transactions = transactions[[
        '户名', '银行名称', '开户银行机构名称', '帐号', '交易日期', '交易方式', '资金收付标志', '金额', '余额',
        '交易对手姓名', '交易对手帐号', '交易对手金融机构名称'
    ]]

    return accounts, transactions


# %%
def format_bhyh(dir_path: pathlib.Path):
    format_progress('开始分析渤海银行账户……')

    if dir_path.is_file():
        raise Exception('渤海银行账户应该是一个文件夹！')

    acc_file = next(dir_path.glob('*账户*.*'))
    format_progress(acc_file.name)
    accounts = pd.read_excel(
        acc_file,
        header=1,
        converters={'开户日期': lambda x: pd.to_datetime(str(x))})

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
