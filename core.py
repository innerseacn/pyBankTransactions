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


# 解析流水文件，将结果保存在tmp_trans_list_by_file中，并返回总行数
def parse_trans_file(trans_file: pathlib.Path, bank_para: st.BankPara,
                     tmp_trans_list_by_file: list) -> int:
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
            # 如果本文件名符合如下规则, 此时认为工作表名就是户名
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
        # 如果流水中不包含户名列，则此项不为空，此时使用文件名或父目录名截取户名
        if '户名' not in tmp_transactions.columns:
            if bank_para.use_dir_name:
                tmp_name = trans_file.parent.stem
            else:
                tmp_name = trans_file.stem
            tmp_transactions['户名'] = get_account_name(tmp_name,
                                                      bank_para.deco_strings)
        tmp_trans_list_by_file.append(tmp_transactions)
    except KeyError:
        format_progress('工作表映射错误，跳过文件\n                ', True)
    format_progress('工作表解析成功{}/失败{}，解析流水{}/{}条'.format(
        len(tmp_trans_list_by_account), tmp_not_parsed, len(tmp_transactions),
        tmp_line_num))
    return tmp_line_num


def parse_base_dir(dir_path: pathlib.Path,
                   bank_para: st.BankPara) -> pd.DataFrame:
    format_progress('开始分析{}账户……'.format(dir_path.name))
    tmp_trans_list_by_file = []  # 流水列表（按文件）
    tmp_all_nums = 0  # 所有流水行数
    for trans_file in dir_path.iterdir():
        if trans_file.is_dir():  # 对每一个子目录
            format_progress('  进入子目录——{}……'.format(trans_file.name))
            for sub_file in trans_file.iterdir():  # 解析子目录所有文件
                tmp_all_nums += parse_trans_file(sub_file, bank_para,
                                                 tmp_trans_list_by_file)
        else:  # 对每一个文件
            if trans_file.match('~*') or trans_file.match('*账户情况.xls*'):
                continue  # 如是无用文件则跳过
            else:
                tmp_all_nums += parse_trans_file(trans_file, bank_para,
                                                 tmp_trans_list_by_file)
    tmp_trans = pd.concat(tmp_trans_list_by_file, sort=False)
    tmp_trans['银行名称'] = dir_path.name
    if bank_para.second_amount_col is not None:
        combine_amount_cols(tmp_trans, bank_para.second_amount_col)
    # tmp_trans = tmp_trans.reindex(columns=st.COLUMN_ORDER)
    tmp_trans['交易日期'] = pd.to_datetime(tmp_trans['交易日期'])
    tmp_trans['交易金额'] = pd.to_numeric(tmp_trans['交易金额'])
    if not bank_para.has_minus_amounts:
        tmp_trans['借贷标志'] = tmp_trans['借贷标志'].str.strip()
        charge_off_amount(tmp_trans['交易金额'], tmp_trans['借贷标志'])
    format_progress('    分析结束，共解析{}/{}条'.format(len(tmp_trans), tmp_all_nums))
    if len(tmp_trans) < tmp_all_nums:
        format_progress(
            '                   ^----------------------------- 请查找问题，若无问题请忽略')
    return tmp_trans
