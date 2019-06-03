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


# 在两列金额中，有空值则返回空值；没有空值则返回0值
def get_none_or_zero_lines(amount_col: pd.Series) -> pd.Series:
    num = pd.to_numeric(amount_col)
    num_na = num.isna()
    if num_na.any():
        return num_na
    else:
        return num == 0


# 将出账金额置为复数，方便观看。但要注意有些冲抵金额本身为负，
# 应用本函数后变为正数，所以自动化计算需要结合收付标志处理
def amount_set_minus(trans: pd.DataFrame,
                     second_amount_col: str = None) -> None:
    if '借贷标志' in trans.columns:
        trans['交易金额'][trans['借贷标志'].isin(st.CHARGE_OFF_WORDS)] *= -1
    elif second_amount_col in trans.columns:
        none_or_zero_lines = get_none_or_zero_lines(trans[second_amount_col])
        trans['交易金额'][none_or_zero_lines] *= -1


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
    none_or_zero_lines = get_none_or_zero_lines(data['交易金额'])
    data['交易金额'][none_or_zero_lines] = data[second_amount_col]


# 以下是特殊解析方法
# 中国银行
def parse_trans_boc(excel_file: pd.ExcelFile, tmp_trans_list_by_sheet) -> int:
    col_map = {
        '姓名': '户名',
        '客户姓名': '户名',
        '交易后可抵用金额': '账户余额',
        '主账号': '账号',
        '货币': '币种',
        '交易类型': '交易方式',
        '交易发生日': '交易日期',
        '柜员和分行': '柜员号',
        '借贷': '借贷标志',
        '交易后金额': '账户余额',
        '对方姓名': '对方户名',
        '主账户账号': '账号',
        '交易类型描述': '交易方式',
        '交易柜员': '柜员号',
        '交易机构名称': '交易网点',
        '交易码': '交易代码',
        '交易货币': '币种',
        '借贷标识': '借贷标志',
        '交易后余额': '账户余额',
        '对方账户名': '对方户名',
    }
    tmp_line_num = 0
    # 解析旧线
    if '旧线账号' in excel_file.sheet_names:
        old_line_accs = excel_file.parse(sheet_name='旧线账号',
                                         usecols='C,D',
                                         dtype=str)
    elif '旧账号' in excel_file.sheet_names:
        old_line_accs = excel_file.parse(sheet_name='旧账号',
                                         usecols='C,D',
                                         dtype=str)
    else:
        format_progress(excel_file.sheet_names)
        raise Exception('中国银行不包含旧线账号或旧账号')
    old_line_accs.dropna(axis=0, how='any', subset=['卡号'], inplace=True)
    old_line_accs.drop_duplicates(inplace=True)
    old_line_accs['卡号'] = old_line_accs.groupby('账号')['卡号'].apply(
        '/'.join).reset_index()
    old_line_trans = excel_file.parse(sheet_name='旧线交易',
                                      usecols='A:M',
                                      dtype=str)
    tmp_line_num += len(old_line_trans)
    old_line_trans = pd.merge(old_line_trans,
                              old_line_accs,
                              how='left',
                              on='账号',
                              validate='m:1')
    old_line_trans.rename(columns=col_map, inplace=True)
    tmp_trans_list_by_sheet.append(old_line_trans)

    # 解析新线
    new_line_accs = excel_file.parse(sheet_name='新线账号',
                                     usecols='A,D,E',
                                     dtype=str)
    new_line_accs.drop_duplicates(inplace=True)
    new_line_accs_no_na = new_line_accs.dropna(axis=0,
                                               how='any',
                                               subset=['卡号'])
    # 存在同一账号对应多个卡号的情况，将其合并到一行
    new_line_accs_no_na = new_line_accs_no_na.groupby(
        ['姓名', '主账号'])['卡号'].apply('/'.join).reset_index()
    new_line_trans = excel_file.parse(sheet_name='新线交易', dtype=str)
    tmp_line_num += len(new_line_trans)
    new_line_trans = pd.merge(new_line_trans,
                              new_line_accs_no_na[['卡号', '主账号']],
                              how='left',
                              on='主账号',
                              validate='m:1')
    new_line_trans.rename(columns=col_map, inplace=True)
    tmp_trans_list_by_sheet.append(new_line_trans)

    # 解析20150701后交易
    newer_trans = excel_file.parse(sheet_name='20150701后交易',
                                   usecols='D:G,I:AE',
                                   dtype=str)
    tmp_line_num += len(newer_trans)
    newer_trans = pd.merge(newer_trans,
                           new_line_accs[['姓名', '主账号'
                                          ]].rename(columns={'主账号': '主账户账号'}),
                           how='left',
                           on='主账户账号',
                           validate='m:1')
    newer_trans.rename(columns=col_map, inplace=True)
    tmp_trans_list_by_sheet.append(newer_trans)
    return tmp_line_num


# 建设银行
def parse_trans_ccb(excel_file: pd.ExcelFile, tmp_trans_list_by_sheet) -> int:
    def _parse_sheet(row_data, header, trans_list_by_sheet):
        header_lines = row_data.index[row_data.duplicated()].to_list()
        header_lines.append(len(row_data) + 1)
        _begin = 2
        line_num = 0
        for _end in header_lines:
            tmp_trans = row_data.iloc[_begin:_end - 1]
            tmp_trans.columns = header.columns
            tmp_acc_str = row_data.iloc[_begin - 2, 0]
            if isinstance(tmp_acc_str, str):
                tmp_acc = tmp_acc_str.replace('，', ':').split(':')
                tmp_trans['户名'] = tmp_acc[1]
                tmp_trans['账号'] = tmp_acc[5]
                tmp_trans['币种'] = tmp_acc[9]
            if tmp_trans.iloc[0][0] != '查无结果':
                line_num += len(tmp_trans)
            trans_list_by_sheet.append(tmp_trans)
            _begin = _end + 1
        return line_num

    col_map = {
        '交易卡号': '卡号',
        '借方发生额': '交易金额',
        '交易渠道': '交易方式',
        '交易机构名称': '交易网点',
        '对方行名': '对方开户行',
        '交易备注': '备注',
        '借贷方标志': '借贷标志',
    }
    tmp_line_num = 0
    # 分析活期流水
    header1 = excel_file.parse(sheet_name='个人活期明细信息-新一代', header=9, nrows=0)
    header1.rename(columns=col_map, inplace=True)
    row_data1 = excel_file.parse(sheet_name='个人活期明细信息-新一代',
                                 header=None,
                                 skiprows=8,
                                 dtype=str)
    first_amount_col = pd.to_numeric(row_data1[5], errors='coerce') * -1
    second_amount_col = pd.to_numeric(row_data1[6], errors='coerce')
    first_amount_col[first_amount_col == 0] = second_amount_col
    row_data1[5] = first_amount_col
    tmp_line_num += _parse_sheet(row_data1, header1, tmp_trans_list_by_sheet)
    # 分析定期流水
    header2 = excel_file.parse(sheet_name='个人定期明细信息-新一代', header=9, nrows=0)
    header2.rename(columns=col_map, inplace=True)
    row_data2 = excel_file.parse(sheet_name='个人定期明细信息-新一代',
                                 header=None,
                                 skiprows=8,
                                 dtype=str)
    tmp_line_num += _parse_sheet(row_data2, header2, tmp_trans_list_by_sheet)
    return tmp_line_num


# 邮储银行
def parse_trans_psbc(excel_file: pd.ExcelFile, tmp_trans_list_by_sheet) -> int:
    col_map = {
        '交易渠道': '交易方式',
        '交易机构名称': '交易网点',
        '对方账号/卡号/汇票号': '对方账号',
        '对方开户机构': '对方开户行',
    }
    tmp_line_num = 0
    for sheet in excel_file.sheet_names:
        tmp_acc_strs = excel_file.parse(sheet_name=sheet, header=None, nrows=4)
        _tmp_str = tmp_acc_strs.iloc[1, 0].split(':')
        _name = _tmp_str[2]
        _account = _tmp_str[1].split()[0]
        _currency = tmp_acc_strs.iloc[3, 0].split(':')[1].split()[0]
        tmp_trans_sheet = excel_file.parse(sheet_name=sheet,
                                           header=5,
                                           dtype=str,
                                           skipfooter=3)
        tmp_trans_sheet.columns = tmp_trans_sheet.columns.str.strip()
        tmp_trans_sheet.rename(columns=col_map, inplace=True)
        tmp_trans_sheet['户名'] = _name
        tmp_trans_sheet['账号'] = _account
        tmp_trans_sheet['币种'] = _currency
        tmp_trans_list_by_sheet.append(tmp_trans_sheet)
        tmp_line_num += len(tmp_trans_sheet)
    return tmp_line_num


# 解析流水文件，将结果保存在tmp_trans_list_by_file中，并返回总行数
def parse_trans_file(trans_file: pathlib.Path, bank_para: st.BankPara,
                     tmp_trans_list_by_file: list) -> int:
    format_progress('    {}……'.format(trans_file.name), True)
    excel_file = pd.ExcelFile(trans_file)
    tmp_trans_list_by_sheet = []  # 当前文件流水列表（按工作表）
    tmp_not_parsed = 0  # 当前文件无法解析工作表数
    tmp_line_num = 0  # 当前文件流水行数
    if bank_para.special_func == '中国银行':
        tmp_line_num = parse_trans_boc(excel_file, tmp_trans_list_by_sheet)
    elif bank_para.special_func == '建设银行':
        tmp_line_num = parse_trans_ccb(excel_file, tmp_trans_list_by_sheet)
    elif bank_para.special_func == '邮储银行':
        tmp_line_num = parse_trans_psbc(excel_file, tmp_trans_list_by_sheet)
    else:  # 其他银行
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
                if bank_para.sheet_name_is == '户名':
                    tmp_trans_sheet['户名'] = sheet
                elif bank_para.sheet_name_is == '账号':
                    tmp_trans_sheet['账号'] = sheet
                tmp_trans_list_by_sheet.append(tmp_trans_sheet)
                if tmp_trans_sheet.iloc[0][0] not in st.NONE_TRANS_WORDS:
                    tmp_line_num += len(tmp_trans_sheet)
                continue
    tmp_transactions = pd.concat(tmp_trans_list_by_sheet,
                                 ignore_index=True,
                                 sort=False)
    try:
        if bank_para.second_amount_col is not None:
            combine_amount_cols(tmp_transactions, bank_para.second_amount_col)
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
        if (tmp_line_num -
                len(tmp_transactions)) == (bank_para.footer *
                                           len(tmp_trans_list_by_sheet)):
            tmp_line_num = len(tmp_transactions)
    except KeyError as k:
        format_progress('字段{}映射错误，跳过文件\n                '.format(k), True)
    format_progress('工作表解析成功{}/失败{}，解析流水{}/{}条'.format(
        len(tmp_trans_list_by_sheet), tmp_not_parsed, len(tmp_transactions),
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
            if trans_file.match('~*') or trans_file.match(dir_path.name +
                                                          '账户*'):
                continue  # 如是无用文件则跳过
            else:
                tmp_all_nums += parse_trans_file(trans_file, bank_para,
                                                 tmp_trans_list_by_file)
    tmp_trans = pd.concat(tmp_trans_list_by_file,
                          ignore_index=True,
                          sort=False)
    tmp_trans['银行名称'] = dir_path.name
    # tmp_trans = tmp_trans.reindex(columns=st.COLUMN_ORDER)
    tmp_trans['交易日期'] = pd.to_datetime(tmp_trans['交易日期'], errors='coerce')
    tmp_trans['交易金额'] = pd.to_numeric(tmp_trans['交易金额'])
    if not bank_para.has_minus_amounts:
        amount_set_minus(tmp_trans,
                         second_amount_col=bank_para.second_amount_col)
    format_progress('    分析结束，共解析{}/{}条'.format(len(tmp_trans), tmp_all_nums))
    if len(tmp_trans) < tmp_all_nums:
        format_progress(
            '                   ^----------------------------- 请查找问题，若无问题请忽略')
    return tmp_trans


def format_transactions(base_path: pathlib.Path) -> pd.DataFrame:
    format_progress('开始分析银行流水……')
    tmp_trans_list_by_bank = []
    tmp_banks_no_support = 0
    for dir in base_path.iterdir():
        try:
            if dir.is_dir():
                tmp_trans_list_by_bank.append(
                    parse_base_dir(dir, st.BANK_PARAS[dir.name]))
        except KeyError as k:
            tmp_banks_no_support += 1
            format_progress('暂不支持{}'.format(k))
    transactions = pd.concat(tmp_trans_list_by_bank,
                             ignore_index=True,
                             sort=False)
    transactions = transactions.reindex(columns=st.COLUMN_ORDER)
    tmp_cols = transactions.select_dtypes(include='object').columns
    for col in tmp_cols:
        transactions[col] = transactions[col].str.strip()
    transactions.sort_values(by='交易日期', inplace=True)
    format_progress('全部分析完成，\n    成功解析银行{}家，流水{}条\n    发现暂不支持银行{}家'.format(
        len(tmp_trans_list_by_bank), len(transactions), tmp_banks_no_support))
    return transactions
