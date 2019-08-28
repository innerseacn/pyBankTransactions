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


def format_error(msg: str) -> None:
    print('\n   ✘════' + msg + '════', end='')


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
        charge_off_lines = trans['借贷标志'].str.strip().isin(st.CHARGE_OFF_WORDS)
        trans.loc[charge_off_lines, '交易金额'] *= -1
    elif second_amount_col in trans.columns:
        none_or_zero_lines = get_none_or_zero_lines(trans[second_amount_col])
        trans.loc[none_or_zero_lines, '交易金额'] *= -1
    elif '账户余额' in trans.columns:
        less_balance_lines = pd.to_numeric(trans['账户余额']).diff() < 0
        trans.loc[less_balance_lines, '交易金额'] *= -1


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
    data.loc[none_or_zero_lines, '交易金额'] = data[second_amount_col]


# 以下是特殊解析方法
# 中国银行
def parse_trans_boc(excel_file: pd.ExcelFile, tmp_trans_list_by_sheet) -> int:
    col_map = {
        '姓名': '户名',
        '客户姓名': '户名',
        '交易后可抵用金额': '账户余额',
        '子账号': '账号',
        '货币': '币种',
        '交易类型': '交易方式',
        '交易发生日': '交易日期',
        '柜员和分行': '柜员号',
        '借贷': '借贷标志',
        '交易后金额': '账户余额',
        '对方姓名': '对方户名',
        '交易账号': '账号',
        '交易类型描述': '交易方式',
        '交易柜员': '柜员号',
        '交易机构名称': '交易网点',
        '交易码': '交易代码',
        '交易货币': '币种',
        '借贷方向': '借贷标志',
        '借贷标识': '借贷标志',
        '交易后余额': '账户余额',
        '对方账户名': '对方户名',
    }
    tmp_line_num = 0

    # 解析新线
    if '新线交易' in excel_file.sheet_names:
        if '新线账号' in excel_file.sheet_names:
            new_line_accs = excel_file.parse(sheet_name='新线账号',
                                             usecols='A,D,F,I',
                                             dtype=str)
            new_line_accs.dropna(axis=0,
                                 how='any',
                                 subset=['子账号'],
                                 inplace=True)
            new_line_accs.drop_duplicates(inplace=True)
            new_line_accs_no_na = new_line_accs.dropna(axis=0,
                                                       how='any',
                                                       subset=['卡号'])
            # 存在同一账号对应多个卡号的情况，将其合并到一行
            new_line_accs_no_na_group = new_line_accs_no_na.groupby(
                ['姓名', '子账号'])['卡号'].apply('/'.join).reset_index()
        else:
            format_error('本文件不包含新线账号')
        new_line_trans = excel_file.parse(sheet_name='新线交易', dtype=str)
        tmp_line_num += len(new_line_trans)
        new_line_trans = pd.merge(new_line_trans,
                                  new_line_accs_no_na_group[['卡号', '子账号']],
                                  how='left',
                                  on='子账号',
                                  validate='m:1')
        new_line_trans.rename(columns=col_map, inplace=True)
        tmp_trans_list_by_sheet.append(new_line_trans)
    elif '新线流水' in excel_file.sheet_names:
        new_line_trans = excel_file.parse(sheet_name='新线流水', dtype=str)
        tmp_line_num += len(new_line_trans)
        new_line_trans.rename(columns=col_map, inplace=True)
        tmp_trans_list_by_sheet.append(new_line_trans)
    else:
        format_error('本文件不包含新线交易或新线流水')

    # 解析旧线
    if '旧线交易' in excel_file.sheet_names:
        if '旧线账号' in excel_file.sheet_names:
            old_line_accs = excel_file.parse(sheet_name='旧线账号',
                                             usecols='C,D',
                                             dtype=str)
        elif '旧账号' in excel_file.sheet_names:
            old_line_accs = excel_file.parse(sheet_name='旧账号',
                                             usecols='C,D',
                                             dtype=str)
        elif '旧账号' in new_line_accs_no_na.columns:
            old_line_accs = new_line_accs_no_na[['卡号', '旧账号']].copy()
            old_line_accs.columns = ['卡号', '账号']
        else:
            format_error('本文件不包含旧线账号或旧账号')
        old_line_accs.dropna(axis=0, how='any', subset=['卡号'], inplace=True)
        old_line_accs.drop_duplicates(inplace=True)
        old_line_accs = old_line_accs.groupby('账号')['卡号'].apply(
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
    else:
        format_error('本文件不包含旧线交易')

    # 解析20150701后交易
    if '20150701后交易' in excel_file.sheet_names:
        newer_trans = excel_file.parse(sheet_name='20150701后交易',
                                       usecols='A:G,I:AE',
                                       dtype=str)
        tmp_line_num += len(newer_trans)
        if pd.isna(newer_trans.loc[0, '姓名']):
            del newer_trans['姓名']
            new_line_accs_name = new_line_accs[['姓名',
                                                '子账号']].rename(columns={
                                                    '子账号': '交易账号'
                                                }).drop_duplicates()
            newer_trans = pd.merge(newer_trans,
                                   new_line_accs_name,
                                   how='left',
                                   on='交易账号',
                                   validate='m:1')
        newer_trans.rename(columns=col_map, inplace=True)
        tmp_trans_list_by_sheet.append(newer_trans)
    elif '20120720后交易流水' in excel_file.sheet_names:
        newer_trans = excel_file.parse(sheet_name='20120720后交易流水',
                                       usecols='A:G,I:AE',
                                       dtype=str)
        tmp_line_num += len(newer_trans)
        newer_trans.rename(columns=col_map, inplace=True)
        tmp_trans_list_by_sheet.append(newer_trans)
    else:
        format_error('本文件不包含20120720后交易流水')

    # 去除账号中的先导0
    # for _df in tmp_trans_list_by_sheet:
    #     _df['账号'] = pd.to_numeric(_df['账号'], errors='ignore').apply(str)
    return tmp_line_num


# 建设银行
def parse_trans_ccb(excel_file: pd.ExcelFile, tmp_trans_list_by_sheet) -> int:
    def _parse_sheet(row_data, header, trans_list_by_sheet):
        header_lines = row_data.index[row_data.duplicated()].to_list()
        header_lines.append(len(row_data) + 1)
        _begin = 2
        line_num = 0
        for _end in header_lines:
            tmp_trans = row_data.iloc[_begin:_end - 1].copy()
            tmp_trans.columns = header.columns
            tmp_acc_str = row_data.iloc[_begin - 2, 0]
            if isinstance(tmp_acc_str, str):
                tmp_acc = tmp_acc_str.replace('，', ':').split(':')
                tmp_trans['户名'] = tmp_acc[1]
                tmp_trans['账号'] = tmp_acc[5]
                tmp_trans['币种'] = tmp_acc[9]
            if tmp_trans.iloc[0, 0] != '查无结果':
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
    fixed_deposit_str = ''
    current_deposit_str = ''
    # 判断是个人还是企业
    if '个人活期明细信息-新一代' in excel_file.sheet_names:
        current_deposit_str = '个人活期明细信息-新一代'
        fixed_deposit_str = '个人定期明细信息-新一代'
    else:
        current_deposit_str = '企业活期明细信息'
        fixed_deposit_str = '企业定期明细信息'

    # 分析活期流水
    header1 = excel_file.parse(sheet_name=current_deposit_str, header=9, nrows=0)
    header1.rename(columns=col_map, inplace=True)
    row_data1 = excel_file.parse(sheet_name=current_deposit_str,
                                 header=None,
                                 skiprows=8,
                                 dtype=str)
    first_amount_col = pd.to_numeric(row_data1[5], errors='coerce') * -1
    second_amount_col = pd.to_numeric(row_data1[6], errors='coerce')
    first_amount_col[first_amount_col == 0] = second_amount_col
    row_data1[5] = first_amount_col
    tmp_line_num += _parse_sheet(row_data1, header1, tmp_trans_list_by_sheet)
    # 分析定期流水
    header2 = excel_file.parse(sheet_name=fixed_deposit_str, header=9, nrows=0)
    header2.rename(columns=col_map, inplace=True)
    row_data2 = excel_file.parse(sheet_name=fixed_deposit_str,
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
        if len(tmp_acc_strs) == 0:
            continue
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


# 平安银行
def parse_trans_pab(excel_file: pd.ExcelFile, tmp_trans_list_by_sheet) -> int:
    col_map = {
        '借方发生额': '交易金额',
        '交易对方户名': '对方户名',
        '交易对方账号': '对方账号',
        '交易对方行名称': '对方开户行',
    }
    tmp_line_num = 0
    for sheet in excel_file.sheet_names:
        tmp_acc_strs = excel_file.parse(sheet_name=sheet, header=None, nrows=5)
        tmp_acc_strs.dropna(how='all',axis=1,inplace=True)
        _name = tmp_acc_strs.iloc[1, 3]
        _account = tmp_acc_strs.iloc[1, 1]
        _card_num = tmp_acc_strs.iloc[2, 1]
        _currency = tmp_acc_strs.iloc[4, 3]
        tmp_trans_sheet = excel_file.parse(sheet_name=sheet,
                                           header=6,
                                           dtype=str,
                                           skipfooter=2)
        tmp_trans_sheet.rename(columns=col_map, inplace=True)
        tmp_trans_sheet['户名'] = _name
        tmp_trans_sheet['账号'] = _account
        tmp_trans_sheet['卡号'] = _card_num
        tmp_trans_sheet['币种'] = _currency
        tmp_trans_sheet['交易金额'] = tmp_trans_sheet['交易金额'].str.replace(',', '')
        tmp_trans_sheet['贷方发生额'] = tmp_trans_sheet['贷方发生额'].str.replace(
            ',', '')
        tmp_trans_list_by_sheet.append(tmp_trans_sheet)
        tmp_line_num += len(tmp_trans_sheet)
    return tmp_line_num


# 华夏银行
def parse_trans_hxb(excel_file: pd.ExcelFile, tmp_trans_list_by_sheet) -> int:
    col_map = {
        '客户名称': '户名',
        '过账日期': '交易日期',
        '业务类型': '交易方式',
        '发生额': '交易金额',
        '余额': '账户余额',
        '凭证号': '卡号',
        '对方户名(或商户名称)': '对方户名',
        '对方账号(或商户编号)': '对方账号',
        '对方银行': '对方开户行'
    }
    tmp_line_num = 0
    for sheet in excel_file.sheet_names:
        header = get_header(excel_file, sheet, st.TEST_HEADER)  # 寻找表头
        if header == 0:
            tmp_trans_sheet = excel_file.parse(sheet_name=sheet,
                                               header=header,
                                               dtype=str)
            tmp_trans_sheet.rename(columns=col_map, inplace=True)
        elif header == 2:
            tmp_acc_strs = excel_file.parse(sheet_name=sheet,
                                            header=None,
                                            nrows=2)
            _tmp_str = tmp_acc_strs.iloc[1, 0].split('：')
            _name = _tmp_str[4]
            _account = _tmp_str[1].split()[0]
            _card_num = _tmp_str[3].split()[0]
            tmp_trans_sheet = excel_file.parse(sheet_name=sheet,
                                               header=2,
                                               dtype=str)
            tmp_trans_sheet.rename(columns=col_map, inplace=True)
            tmp_trans_sheet['户名'] = _name
            tmp_trans_sheet['账号'] = _account
            tmp_trans_sheet['卡号'] = _card_num
        elif header == -1:
            continue
        else:
            format_error('{}无法解析，跳过'.format(sheet))
            continue
        tmp_trans_list_by_sheet.append(tmp_trans_sheet)
        tmp_line_num += len(tmp_trans_sheet)
    return tmp_line_num


# 一般银行分析
def parse_trans_common(excel_file: pd.ExcelFile, bank_para: st.BankPara,
                       tmp_trans_list_by_sheet) -> int:
    tmp_line_num = 0
    for sheet in excel_file.sheet_names:  # 对每一个工作表
        header = get_header(excel_file, sheet, st.TEST_HEADER)  # 寻找表头
        if header == -1:  # 空工作表
            continue
        elif header == -2:  # 含数据但表头超过测试数而无法解析的工作表
            format_error('{}无法解析，跳过'.format(sheet))
            continue
        else:  # 找到表头
            tmp_trans_sheet = excel_file.parse(sheet_name=sheet,
                                               header=header,
                                               dtype=str)
            if len(tmp_trans_sheet) == 0:
                continue
            tmp_trans_sheet.rename(columns=bank_para.col_map, inplace=True)
            # 识别非流水表和空数据表
            if '交易日期' not in tmp_trans_sheet.columns:
                if not bank_para.has_nodata_sheets:
                    format_error('{}不包含交易日期，跳过'.format(sheet))
                continue
            if tmp_trans_sheet['交易日期'].isna().all():
                if not bank_para.has_empty_sheets:
                    format_error('{}交易日期不完整，跳过'.format(sheet))
                continue
            # 如果本文件名符合如下规则, 此时认为工作表名就是户名
            if bank_para.sheet_name_is == '户名':
                tmp_trans_sheet['户名'] = sheet
            elif bank_para.sheet_name_is == '账号':
                tmp_trans_sheet['账号'] = sheet
            tmp_trans_list_by_sheet.append(tmp_trans_sheet)
            if tmp_trans_sheet.iloc[0, 0] not in st.NONE_TRANS_WORDS:
                tmp_line_num += len(tmp_trans_sheet)
            continue
    return tmp_line_num


# 解析流水文件，将结果保存在tmp_trans_list_by_file中，并返回总行数
def parse_trans_file(trans_file: pathlib.Path, bank_para: st.BankPara,
                     tmp_trans_list_by_file: list) -> int:
    format_progress('    {}……'.format(trans_file.name), True)
    excel_file = pd.ExcelFile(trans_file)
    tmp_trans_list_by_sheet = []  # 当前文件流水列表（按工作表）
    tmp_line_num = 0  # 当前文件流水行数
    if bank_para.special_func == '中国银行':
        tmp_line_num = parse_trans_boc(excel_file, tmp_trans_list_by_sheet)
    elif bank_para.special_func == '建设银行':
        tmp_line_num = parse_trans_ccb(excel_file, tmp_trans_list_by_sheet)
    elif bank_para.special_func == '邮储银行':
        tmp_line_num = parse_trans_psbc(excel_file, tmp_trans_list_by_sheet)
    elif bank_para.special_func == '平安银行':
        tmp_line_num = parse_trans_pab(excel_file, tmp_trans_list_by_sheet)
    elif bank_para.special_func == '华夏银行':
        tmp_line_num = parse_trans_hxb(excel_file, tmp_trans_list_by_sheet)
    else:  # 其他银行
        tmp_line_num = parse_trans_common(excel_file, bank_para,
                                          tmp_trans_list_by_sheet)
    try:
        tmp_transactions = pd.concat(tmp_trans_list_by_sheet,
                                     ignore_index=True,
                                     sort=False)
        if bank_para.second_amount_col is not None:
            combine_amount_cols(tmp_transactions, bank_para.second_amount_col)
        tmp_transactions.dropna(axis=0, subset=['交易日期', '交易金额'], inplace=True)
        # _tmp_thresh = len(tmp_transactions.columns) / 3
        # tmp_transactions.dropna(axis=0, thresh=_tmp_thresh, inplace=True)
        # 如果流水中不包含户名列，则此项不为空，此时使用文件名或父目录名截取户名
        if ('户名' not in tmp_transactions.columns
            ) or tmp_transactions['户名'].hasnans:
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
    except ValueError as v:
        format_error('无法生成流水，跳过文件（原因：{}）\n   ✘'.format(v))
        # raise v
        tmp_transactions = []
    except KeyError as k:
        format_error('字段{}映射错误，跳过文件'.format(k))
    format_progress('工作表解析成功{}，解析流水{}/{}条'.format(len(tmp_trans_list_by_sheet),
                                                  len(tmp_transactions),
                                                  tmp_line_num))
    return tmp_line_num


def parse_base_dir(dir_path: pathlib.Path,
                   bank_para: st.BankPara) -> (pd.DataFrame, bool):
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
    tmp_trans['交易日期'] = pd.to_datetime(tmp_trans['交易日期'], errors='coerce')
    tmp_trans['交易金额'] = pd.to_numeric(tmp_trans['交易金额'])
    if not bank_para.has_minus_amounts:
        amount_set_minus(tmp_trans,
                         second_amount_col=bank_para.second_amount_col)
    format_progress('    分析结束，共解析{}/{}条'.format(len(tmp_trans), tmp_all_nums))
    # 检测结果正确性
    na_nums = tmp_trans.reindex(columns=bank_para.check_cols).isna().sum()
    no_acc = tmp_trans.reindex(columns=['账号', '卡号']).isna().all(axis=1).sum()
    na_nums.loc['账号和卡号'] = no_acc
    na_nums = na_nums[na_nums > 0]
    na_cols = bank_para.need_cols - set(tmp_trans.columns)
    nag_amounts = len(tmp_trans[tmp_trans['交易金额'] < 0])
    _has_mistakes = False
    if (tmp_trans['户名'].str.find('逐笔明细') != -1).any():
        format_progress('    ✘户名解析不正确。')
        _has_mistakes = True
    if len(tmp_trans) < tmp_all_nums:
        format_progress('    ✘存在未解析数据行。')
        _has_mistakes = True
    if len(na_nums) > 0:
        format_progress('    ✘以下关键字段存在空值：' + str(na_nums.to_dict()))
        _has_mistakes = True
    if len(na_cols) > 0:
        format_progress('    ✘以下所需字段未正确转换：' + str(na_cols))
        _has_mistakes = True
    if nag_amounts == 0:
        format_progress('    ✘交易金额全为正值。')
        _has_mistakes = True
    if _has_mistakes:
        format_progress(
            '✘═══╩════════════════════════════════════════请查找问题，或调整不规范数据！')
    else:
        format_progress('  ✔')
    return tmp_trans, _has_mistakes


def format_transactions(base_path: pathlib.Path) -> pd.DataFrame:
    _num_mistakes = 0
    format_progress('开始分析银行流水……')
    tmp_trans_list_by_bank = []
    tmp_banks_no_support = 0
    for dir in base_path.iterdir():
        try:
            if dir.is_dir():
                _tmp_trans, _has_mistakes = parse_base_dir(
                    dir, st.BANK_PARAS[dir.name])
                if _has_mistakes:
                    _num_mistakes += 1
                tmp_trans_list_by_bank.append(_tmp_trans)
        except KeyError as k:
            tmp_banks_no_support += 1
            format_progress('暂不支持{}'.format(k))
    transactions = pd.concat(tmp_trans_list_by_bank,
                             ignore_index=True,
                             sort=False)
    transactions = transactions.reindex(columns=st.COLUMN_ORDER)
    transactions.dropna(axis=1, how='all', inplace=True)
    tmp_cols = transactions.select_dtypes(include='object').columns
    for col in tmp_cols:
        transactions[col] = transactions[col].str.strip()
    transactions.sort_values(by='交易日期', inplace=True)
    transactions.dropna(axis=1, how='all', inplace=True)
    # 扩展原始列加速分析
    transactions.insert(9, '金额绝对值', transactions['交易金额'].abs())
    format_progress(
        '全部分析完成，\n    成功解析银行{}家，流水{}条\n    存在解析错误银行{}家\n    发现暂不支持银行{}家'.
        format(len(tmp_trans_list_by_bank), len(transactions), _num_mistakes,
               tmp_banks_no_support))
    return transactions


def write_excel(df: pd.DataFrame, path: pathlib.Path) -> None:
    df.to_excel(path / '规范交易流水（张楠制作）.xlsx', index=False, engine='xlsxwriter')
