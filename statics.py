from typing import List, Dict, Union, Set

TEST_HEADER = 3
CHARGE_OFF_WORDS = {'付', '支出', '借', '借方', '出账', '转出', 'D', '0'}
NONE_TRANS_WORDS = {'无交易', '在我行仅有信用卡账户'}
COLUMN_ORDER = [
    '银行名称', '户名', '账号', '卡号', '交易日期', '借贷标志', '币种', '交易金额', '账户余额', '交易方式', 
    '备注', '摘要', '附言', '对方户名', '对方账号', '对方开户行', '交易场所', '交易地区', '交易网点', '柜员号',
    '涉外交易代码', '交易代码', '代办人', '代办人证件'
]
# "银行业金融机构报告可疑交易逐笔明细表"的默认列映射
COL_MAP_COMMON = {
    '资金收付标志': '借贷标志',
    ' 交易额(按原币计)': '交易金额',
    '交易对手姓名或名称': '对方户名',
    '交易对手账号': '对方账号',
    '对方金融机构网点名称': '对方开户行',
    '金融机构名称': '交易网点',
    '涉外收支交易分类与代码': '涉外交易代码',
    '业务标示号': '交易代码',
    '代办人姓名': '代办人',
    '代办人身份证件/证明文件号码': '代办人证件',
    '资金来源和用途': '摘要'
}
CHECK_COLS = {'银行名称', '户名', '借贷标志', '交易日期', '交易金额', '账户余额'}
CHECK_COLS_COMMON = {'银行名称', '户名', '借贷标志', '交易日期', '交易金额'}
CHECK_COLS_NO_SIGN = {'银行名称', '户名', '交易日期', '交易金额', '账户余额'}
NEED_COLS = {'对方户名', '交易网点', '交易方式', '交易代码', '摘要', '备注'}
NEED_COLS_WORDS = {'对方户名', '交易网点', '交易方式', '交易代码', '摘要', '附言', '备注'}
NEED_COLS_NO_REMARKS = {'对方户名', '交易网点', '交易方式', '交易代码', '摘要'}


class BankPara: 
    def __init__(self,
                 col_map: Dict[str, str] = None,
                 sheet_name_is: str = None,
                 has_minus_amounts: bool = False,
                 has_empty_sheets: bool = False,
                 has_nodata_sheets: bool = False,
                 second_amount_col: str = None,
                 deco_strings: Union[str, List[str]] = None,
                 use_dir_name: bool = False,
                 special_func: str = None,
                 skip_files: str = '✔',
                 footer: int = 0,
                 check_cols: Set[str] = CHECK_COLS,
                 need_cols: Set[str] = NEED_COLS) -> None:
        # 列名映射字典{'原列名':'新列名'}，将源文件列名映射到COLUMN_ORDER中的输出列名
        self.col_map = col_map
        # 工作表名的内容是什么（户名、账号）
        self.sheet_name_is = sheet_name_is
        # 金额列支出数额为负数
        self.has_minus_amounts = has_minus_amounts
        # 存在填满空值的空表
        self.has_empty_sheets = has_empty_sheets
        # 存在无效表
        self.has_nodata_sheets = has_nodata_sheets
        # 当流水中入账和出账各自成一列时，此项赋两列的列名，否则为None
        self.second_amount_col = second_amount_col
        # 在包含户名的文件名中将要被截取掉的多余字串，文件中包含户名传入None，否则需要传入需从文件名截取掉的字串，不需截取转入空串
        self.deco_strings = deco_strings
        # 户名实际未包含在文件名中，而是包含在目录名中
        self.use_dir_name = use_dir_name
        # 特殊处理标志，使用专门方法解析
        self.special_func = special_func
        # 本目录下需要跳过的文件，支持通配符
        self.skip_files = skip_files
        # 需要跳过的页脚行数
        self.footer = footer
        # 需要自检的全不为空的字段
        self.check_cols = check_cols
        # 需要自检的包含字段
        self.need_cols = need_cols


BANK_PARAS = {}
BANK_PARAS['北京银行'] = BankPara(
    col_map={
        '帐号': '账号',
        '资金收付标志': '借贷标志',
        '金额': '交易金额',
        '余额': '账户余额',
        '交易对手姓名': '对方户名',
        '交易对手帐号': '对方账号',
        '交易对手金融机构名称': '对方开户行',
        '开户银行机构名称': '交易网点',
        '交易附言': '摘要'
    },
    sheet_name_is='户名',
    footer=1,
    need_cols=(NEED_COLS_NO_REMARKS - {'交易代码'}))
BANK_PARAS['工商银行'] = BankPara(
    col_map={
        '帐号': '账号',
        '服务界面': '交易方式',
        '渠道': '交易方式',
        '发生额': '交易金额',
        '余额': '账户余额',
        '对方帐户': '对方账号',
        '对方开户行名': '对方开户行',
        '对方行名': '对方开户行',
        '交易地区号': '交易地区',
        '交易网点号': '交易网点',
        '注释': '备注',
        '入账日期': '交易日期',
        '入帐日期': '交易日期',
        '交易金额': '金额',
        '对方卡号/账号': '对方账号',
        '对方帐号': '对方账号',
        '对方帐户户名': '对方户名',
        '更新后余额': '账户余额',
        '交易柜员号': '柜员号',
        '交易场所简称': '交易场所',
        '交易描述': '备注',
        '备注1': '备注',
    },
    has_nodata_sheets=True,
    deco_strings='')
BANK_PARAS['广发银行'] = BankPara(
    col_map={
        '客户名称': '户名',
        '本方账号': '账号',
        '本方交易介质': '卡号',
        '交易渠道中文': '交易方式',
        '借贷标识': '借贷标志',
        '交易货币': '币种',
        '当前余额': '账户余额',
        '对手账号名称': '对方户名',
        '对手账号行所号': '对方开户行',
        '交易行': '交易网点',
        '交易柜员': '柜员号',
        '交易码中文': '交易代码',
        '摘要中文': '摘要',
    },
    need_cols=NEED_COLS_WORDS)
BANK_PARAS['哈尔滨银行'] = BankPara(
    col_map={
        '交易时间': '交易日期',
        '渠道名称': '交易方式',
        '借贷标识': '借贷标志',
        '余额': '账户余额',
        '币种名称': '币种',
        '机构号': '交易网点',
        '现转标识': '摘要',
        '附言': '备注',
    },
    footer=5,
    need_cols=(NEED_COLS - {'交易代码', '对方户名'}))
BANK_PARAS['交通银行'] = BankPara(
    col_map={
        '主记账帐号': '账号',
        '主名义账号': '卡号',
        '金额': '交易金额',
        '对方分行': '对方开户行',
        '交易分行': '交易地区',
        '交易网点/部门': '交易网点',
        '交易柜员': '柜员号',
        '业务摘要区': '摘要',
        '帐号': '账号',
        '交易机构所属分行': '交易地区',
        '交易机构号': '交易网点',
        '借贷方标志': '借贷标志',
        '对方帐号': '对方账号',
        '货币码': '币种',
        '技术摘要': '摘要'
    },
    deco_strings='流水',
    check_cols=CHECK_COLS_COMMON,
    need_cols=(NEED_COLS_NO_REMARKS - {'交易代码', '交易方式', '摘要'}))
BANK_PARAS['廊坊银行'] = BankPara(
    col_map={
        '客户账号': '账号',
        '货币代号': '币种',
        '产品说明': '交易方式',
        '借方发生额': '交易金额',
        '对方客户账号': '对方账号',
        '营业机构': '交易网点',
        '柜员代号': '柜员号',
        '代理人姓名': '代办人',
        '代理人证件号码': '代办人证件',
        '摘要描述': '摘要',
    },
    second_amount_col='贷方发生额',
    deco_strings='活期账户流水')
BANK_PARAS['渤海银行'] = BankPara(
    col_map=COL_MAP_COMMON,
    deco_strings='报告可疑交易逐笔明细表—',
    check_cols=CHECK_COLS_COMMON,
    need_cols=(NEED_COLS - {'摘要'}))
BANK_PARAS['光大银行'] = BankPara(
    col_map=COL_MAP_COMMON,
    deco_strings='交易明细',
    check_cols=CHECK_COLS_COMMON)
BANK_PARAS['河北银行'] = BankPara(
    col_map=COL_MAP_COMMON,
    deco_strings='：银行业金融机构报告可疑交易逐笔明细表',
    check_cols=CHECK_COLS_COMMON)
BANK_PARAS['民生银行'] = BankPara(
    col_map={
        '客户姓名': '户名',
        '客户账户': '账号',
        '客户账号': '账号',
        '原币交易金额': '交易金额',
        '对方名称': '对方户名',
        '对方银行名称': '对方开户行'
    },
    check_cols=CHECK_COLS_COMMON,
    need_cols=(NEED_COLS_NO_REMARKS - {'交易代码', '交易网点', '交易方式'}))
BANK_PARAS['浦发银行'] = BankPara(
    col_map={
        '调查户名': '户名',
        '0(支出)/1（收入）': '借贷标志',
        '金额': '交易金额',
        '对方开户银行': '对方开户行',
        '开户银行': '交易网点',
    },
    check_cols=CHECK_COLS_COMMON,
    need_cols=(NEED_COLS - {'交易代码', '摘要'}))
BANK_PARAS['天津农商银行'] = BankPara(
    col_map=COL_MAP_COMMON,
    deco_strings='银行业金融机构报告可疑交易逐笔明细表——',
    check_cols=CHECK_COLS_COMMON)
BANK_PARAS['天津银行'] = BankPara(
    col_map={
        '姓名': '户名',
        '资金收付标志': '借贷标志',
        ' 交易额(按原币计)（元）': '交易金额',
        '交易对手姓名或名称': '对方户名',
        '交易对手账号': '对方账号',
        '对方金融机构网点名称': '对方开户行',
        '金融机构名称': '交易网点',
        '涉外收支交易分类与代码': '涉外交易代码',
        '业务标示号': '交易代码',
        '代办人姓名': '代办人',
        '代办人身份证件/证明文件号码': '代办人证件',
        '资金来源和用途': '摘要'
    },
    has_empty_sheets=True,
    check_cols=CHECK_COLS_COMMON,
    need_cols=NEED_COLS_NO_REMARKS)
BANK_PARAS['兴业银行'] = BankPara(
    col_map={
        '交易机构编号': '交易网点',
        '交易名称': '备注',
        '渠道类型代码': '交易方式',
        '帐号': '账号',
        '对手帐号': '对方账号',
        '对手账号': '对方账号',
        '对手户名': '对方户名',
        '对手开户行': '对方开户行',
        '摘要描述': '摘要',
        '柜员流水号': '柜员号'
    },
    deco_strings='')
BANK_PARAS['渣打银行'] = BankPara(
    col_map=COL_MAP_COMMON,
    deco_strings='',
    footer=1,
    check_cols=CHECK_COLS_COMMON)
BANK_PARAS['中信银行'] = BankPara(
    col_map={
        '交易码': '交易代码',
        '借贷类型代码': '借贷标志',
        '通用对方客户账号': '对方账号',
        '对方账户名称': '对方户名',
        '对方行名称': '对方开户行',
        '客户账号': '账号',
        '客户名称': '户名',
        '核心交易代码': '交易代码',
        '币种中文': '币种',
        '交易用户名': '柜员号',
    },
    need_cols=(NEED_COLS_NO_REMARKS - {'交易网点', '交易方式'}))
BANK_PARAS['招商银行'] = BankPara(
    col_map={
        '客户名称': '户名',
        '交易卡号': '账号',
        '联机余额': '账户余额',
        '交易摘要': '摘要',
        '文字摘要': '备注',
        '对手帐号': '对方账号',
        '对手名称': '对方户名',
        '对手开户行': '对方开户行',
        '我方摘要': '摘要',
        '对方开户机构名称': '对方开户行',
        '对方客户名称': '对方户名',
        '业务编号': '账号',
        '对方业务编号': '对方账号',
        '交易机构': '交易网点',
    },
    has_minus_amounts=True,
    deco_strings='',
    use_dir_name=True,
    check_cols=CHECK_COLS_NO_SIGN,
    need_cols=(NEED_COLS - {'交易代码', '交易方式', '交易网点'}))
BANK_PARAS['农业银行'] = BankPara(
    col_map={
        '合约号': '账号',
        '产品号': '卡号',
        '合约外部服务标识号码': '卡号',
        '合约名称': '户名',
        '借方交易金额': '交易金额',
        '借方金额': '交易金额',
        '贷方金额': '贷方交易金额',
        '交易金额借方': '交易金额',
        '交易金额贷方': '贷方交易金额',
        '贷方交易金额_1': '贷方交易金额',
        '贷方交易金额_': '贷方交易金额',
        '交易后余额': '账户余额',
        '合约账户余额1': '账户余额',
        '合约账户余额': '账户余额',
        '对方银行': '对方开户行',
        '对方开户银行': '对方开户行',
        '交易对手账号': '对方账号',
        '对方名称': '对方户名',
        '交易渠道': '交易方式',
        '渠道代码': '交易方式',
        '摘要信息': '摘要',
        '记账方向标识_1': '借贷标志',
        '交易地点': '交易网点',
        '对方省市代号': '交易地区',
    },
    second_amount_col='贷方交易金额',
    footer=1,
    check_cols=CHECK_COLS_NO_SIGN,
    need_cols=(NEED_COLS_NO_REMARKS - {'交易代码', '交易网点'}))
BANK_PARAS['威海银行'] = BankPara(
    col_map={
        '币别': '币种',
        '借方发生额': '交易金额',
        '交易渠道': '交易方式',
        '交易机构名称': '交易网点',
        '摘要代码': '摘要',
        '对方名称': '对方户名',
        '交易对方行名': '对方开户行',
        '交易类别': '附言',
    },
    sheet_name_is='账号',
    deco_strings='流水',
    second_amount_col='贷方发生额',
    check_cols=CHECK_COLS_NO_SIGN,
    need_cols=(NEED_COLS_WORDS - {'交易代码'}))
BANK_PARAS['中国银行'] = BankPara(
    special_func='中国银行',
    need_cols=NEED_COLS_NO_REMARKS)
BANK_PARAS['建设银行'] = BankPara(
    special_func='建设银行',
    check_cols=CHECK_COLS_NO_SIGN,
    need_cols=(NEED_COLS - {'交易代码'}))
BANK_PARAS['邮储银行'] = BankPara(
    special_func='邮储银行',
    check_cols=CHECK_COLS_NO_SIGN,
    need_cols=(NEED_COLS - {'对方户名', '交易代码', '备注'}))
BANK_PARAS['平安银行'] = BankPara(
    special_func='平安银行',
    second_amount_col='贷方发生额',
    check_cols=CHECK_COLS_NO_SIGN,
    need_cols=(NEED_COLS_NO_REMARKS - {'交易代码', '交易方式'}))
BANK_PARAS['华夏银行'] = BankPara(
    special_func='华夏银行',
    check_cols=CHECK_COLS,
    need_cols=(NEED_COLS_NO_REMARKS - {'交易代码', '交易网点'}))
BANK_PARAS['锦州银行'] = BankPara(
    col_map={
        '交易地点': '交易网点',
        '交易说明': '摘要',
        '支出': '交易金额',
        '交易柜员': '柜员号',
        '余额': '账户余额',
    },
    second_amount_col='存入',
    check_cols=CHECK_COLS_NO_SIGN,
    need_cols={'对方户名', '摘要'})
BANK_PARAS['宁夏银行'] = BankPara(
    special_func='宁夏银行',
    check_cols=CHECK_COLS,
    need_cols=(NEED_COLS_NO_REMARKS - {'交易代码'}))
BANK_PARAS['津南村镇银行'] = BankPara(
    col_map={
        '交易机构': '交易网点',
        '交易名称': '摘要',
        'TRCASH': '交易方式',
        '交易柜员': '柜员号',
        'DRCRIND': '借贷标志',
        '客户名称': '户名',
    },
    skip_files='*大小额明细*',
    check_cols=CHECK_COLS_COMMON,
    need_cols={'交易方式', '摘要', '备注'})
