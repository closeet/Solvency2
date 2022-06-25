import openpyxl
import pickle
from 字段处理 import ws_cell
from 字段处理 import max_row


def serialize(ls, filename):
    with open(filename, 'wb') as file:
        pickle.dump(ls, file)


def deserialize(filename):
    with open(filename, 'rb') as file:
        return pickle.load(file)


dict_deposit_col_name_self = {'购买成本': '存款金额', '认可价值': '认可价值', '应收利息': '应收利息', '存款银行类型': '银行类型', '资产类型': '存款类型'}
ls_deposit_col_name = ['资产简称', '资产全称', '资产大类', '资产类型', '交易对手', '购买成本', '认可价值', '应收利息', '账户', '存款银行类型', '银行资本充足率',
                       '资产五大类分类']
dict_current_deposit_col_name_self = {'资产简称': '帐户信息', '资产全称': '帐户信息', '购买成本': '余额', '认可价值': '余额'}
dict_current_deposit_col_value = {'资产大类': '现金及流动性管理工具', '资产类型': '活期存款', '资产五大类分类': '流动性资产'}

ls_current_deposit = ['资产简称', '资产全称', '资产大类', '资产类型', '交易对手', '购买成本', '认可价值', '账户', '资产五大类分类']

wb_data_rule = openpyxl.load_workbook('数据规则.xlsx')
ws_bank = wb_data_rule['存款交易对手']
dict_bank_counter_party = {}
for row_bank_rule in range(1, ws_bank.max_row+1):
    dict_bank_counter_party[ws_cell(ws_bank, row_bank_rule, 1)] = ws_cell(ws_bank, row_bank_rule, 2)


