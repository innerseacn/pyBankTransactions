# -*- coding: utf-8 -*-
# %%
import pandas as pd
import pathlib
from typing import List

base_path = pathlib.Path(r'C:\Users\InnerSea\工作\反洗钱\附件：罗凯声等人账户交易明细')
test_file = pathlib.Path(base_path / '工商银行/于悦泓.xls')
CHARGE_OFF_WORDS = ['付', '支出', '借', '借方']
TEST_HEADER = 3


# 以下为辅助函数
def format_progress(msg: str):
    print(msg)


def charge_off_amount(amount: pd.Series, sign: pd.Series):
    amount[sign.isin(CHARGE_OFF_WORDS)] *= -1


def get_account_name(name: str, deco_strings: str or List[str] = '') -> str:
    if isinstance(deco_strings, str):
        name = name.replace(deco_strings, '')
    elif isinstance(deco_strings, List[str]):
        for string in deco_strings:
            name = name.replace(string, '')
    return name


# 以下为账户分析函数
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
    excel_file = pd.ExcelFile(map_file)
    map_list = excel_file.parse(
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


# 以下为流水分析函数
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
    transactions['银行名称'] = dir_path.name
    charge_off_amount(transactions['金额'], transactions['资金收付标志'])
    transactions = transactions[[
        '银行名称', '户名', '帐号', '交易日期', '交易方式', '资金收付标志', '金额', '余额', '交易对手姓名',
        '交易对手帐号', '交易对手金融机构名称', '交易附言'
    ]]
    transactions.columns = [
        '银行名称', '户名', '账号', '交易日期', '交易方式', '收付标志', '金额(原币)', '余额', '对手户名',
        '对手账号', '对手开户行', '备注'
    ]
    return transactions


def format_common(dir_path: pathlib.Path,
                  deco_strings: str or List[str] = '') -> pd.DataFrame:
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
    charge_off_amount(transactions[' 交易额(按原币计)'], transactions['资金收付标志'])
    charge_off_amount(transactions[' 交易额(折合美元)'], transactions['资金收付标志'])
    transactions = transactions[[
        '银行名称', '户名', '账号', '交易日期', '交易方式', '资金收付标志', '币种', ' 交易额(按原币计)',
        ' 交易额(折合美元)', '交易对手姓名或名称', '交易对手账号', '涉外收支交易分类与代码', '业务标示号', '代办人姓名',
        '代办人身份证件/证明文件号码', '备注'
    ]]
    transactions.columns = [
        '银行名称', '户名', '账号', '交易日期', '交易方式', '收付标志', '币种', '金额(原币)', '金额(美元)',
        '对手户名', '对手账号', '涉外交易代码', '交易代码', '代办人', '代办人证件', '备注'
    ]
    return transactions


def format_gsyh(dir_path: pathlib.Path) -> pd.DataFrame:
    format_progress('开始分析工商银行账户……')
    tmp_trans_list_by_name = []
    for trans_file in dir_path.iterdir():  # 对每一个excel文件
        if trans_file.match('~*') or trans_file.match('*账户情况.xls*'):
            continue
        format_progress(trans_file.name)
        excel_file = pd.ExcelFile(trans_file)
        tmp_trans_list_by_account = []
        for sheet in excel_file.sheet_names:  # 对每一个工作表
            for header in range(0, TEST_HEADER):  # 尝试解析标题行TEST_HEADER次
                tmp_trans_sheet = excel_file.parse(
                    sheet_name=sheet, nrows=0, header=header)
                if len(tmp_trans_sheet.columns) == 0:
                    break
                if 'Unnamed: 1' not in tmp_trans_sheet.columns:  # 一旦找到标题行则解析
                    tmp_trans_sheet = excel_file.parse(
                        sheet_name=sheet,
                        header=header,
                        dtype={
                            '账号': str,
                            '卡号': str,
                            '交易日期': str,
                            '对方卡号/账号': str,
                            '对方帐户': str,
                            '入账日期': str
                        })
                    tmp_trans_sheet.rename(
                        columns={
                            '入账日期': '交易日期',
                            '对方卡号/账号': '对方帐户',
                            '对方帐户户名': '对方户名',
                            '更新后余额': '余额',
                            '交易柜员号': '柜员号',
                            '交易场所简称': '交易场所',
                            '交易描述': '注释'
                        },
                        inplace=True)
                    tmp_trans_list_by_account.append(tmp_trans_sheet)
                    break
            else:  # 尝试TEST_HEADER次仍未找到标题行则放弃
                format_progress('在前' + TEST_HEADER + '行没找到标题行，无法解析' + sheet)
        tmp_transactions = pd.concat(tmp_trans_list_by_account, sort=False)
        tmp_transactions['户名'] = trans_file.stem
        tmp_trans_list_by_name.append(tmp_transactions)
    transactions = pd.concat(tmp_trans_list_by_name, sort=False)
    transactions['银行名称'] = dir_path.name
    charge_off_amount(transactions['发生额'], transactions['借贷标志'])
    transactions = transactions[[
        '银行名称', '户名', '账号', '卡号', '交易日期', '服务界面', '借贷标志', '币种', '发生额', '余额',
        '对方户名', '对方帐户', '对方开户行名', '交易地区号', '交易场所', '交易网点号', '柜员号', '交易代码', '注释'
    ]]
    transactions.columns = [
        '银行名称', '户名', '账号', '卡号', '交易日期', '交易方式', '收付标志', '币种', '金额(原币)', '余额',
        '对手户名', '对手账号', '对手开户行', '交易地区', '交易场所', '交易网点', '柜员号', '交易代码', '备注'
    ]
    return transactions


# words = [
#     '银行名称', '户名', '账号', '交易日期', '交易方式', '收付标志', '币种', '金额(原币)', '金额(美元)', '余额',
#     '对手户名', '对手账号', '对手开户行', '交易地区', '交易场所', '交易网点', '柜员号', '涉外交易代码', '交易代码', '代办人', '代办人证件', '备注'
# ]
# a = format_gsyh(base_path / '工商银行')


def format_transactions(base_path: pathlib.Path) -> pd.DataFrame:
    format_progress('开始分析银行流水……')
    tmp_trans_list_by_bank = []
    try:
        for dir in base_path.iterdir():
            if dir.is_dir():
                if dir.name == '北京银行':
                    tmp_trans_list_by_bank.append(format_bjyh(dir))
                elif dir.name == '工商银行':
                    tmp_trans_list_by_bank.append(format_gsyh(dir))
                elif dir.name == '渤海银行':
                    tmp_trans_list_by_bank.append(
                        format_common(dir, '报告可疑交易逐笔明细表—'))
                else:
                    pass
        transactions = pd.concat(
            tmp_trans_list_by_bank, ignore_index=True, sort=False)
        transactions.dropna(axis=0, how='any', subset=['交易日期'], inplace=True)
    except Exception as e:
        raise e
    return transactions


# %%
a = format_transactions(base_path)
