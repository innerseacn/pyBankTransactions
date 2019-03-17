# -*- coding: utf-8 -*-
# %%
import pandas as pd
import pathlib

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
    format_progress(doc)
    accounts = pd.read_excel(dir_path / '北京银行账户情况.xlsx', header=1)
    tmp_trans_list = pd.read_excel(
        dir_path / '交易明细模板-罗扬等（京行天分）.xls',
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
        '户名', '交易对手姓名', '交易对手帐号', '交易对手金融机构名称'
    ]]
    format_progress('end 北京银行')
    return accounts, transactions


accounts, transactions = format_bjyh(
    pathlib.Path(r'C:\Users\InnerSea\工作\反洗钱\附件：罗凯声等人账户交易明细\北京银行'))

# %%
tmp = pd.read_excel(
    r"C:\Users\InnerSea\工作\反洗钱\附件：罗凯声等人账户交易明细\北京银行\交易明细模板-罗扬等（京行天分）.xls",
    sheet_name='孔维军')

# %%
accounts
# %%
transactions

# %%
account_list.append(accounts)
transaction_list.append(transactions)
