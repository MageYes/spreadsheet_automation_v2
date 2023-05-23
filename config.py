import pandas as pd
import copy
start_date = pd.Timestamp('2023-05-01')
end_date = pd.Timestamp('2028-12-31')
excel_path = r"D:\GitHub\spreadsheet_automation_v2\spreadsheet_automation.xlsx"

etfs_lst = ['513010', '513180', '513380', '159920', '159866', '513580', '513030']
etfs_calendar_dict = {
    '513010': ['SSE', 'HKEX'],
    '513180': ['SSE', 'HKEX'],
    '513380': ['SSE', 'HKEX'],
    '159920': ['SSE', 'HKEX'],
    '159866': ['SSE', 'XTKS'],
    '513580': ['SSE', 'HKEX'],
    '513030': ['SSE', 'XETR']
}

general_rules_info = pd.read_excel(excel_path, index_col=0, sheet_name='general rules')
specific_rules_info = pd.read_excel(excel_path, index_col=0, sheet_name='specific rules').dropna(how='all')

general_rules_dict = dict()
for code in general_rules_info.index:
    general_rules_i = general_rules_info.loc[code]
    general_rules_dict[str(code)] = {'Fixing': general_rules_i.loc['Fixing'], 'FX': general_rules_i.loc['FX']}

specific_rules_dict = dict()
for code in specific_rules_info.index:
    specific_rules_i = specific_rules_info.loc[code].dropna()
    specific_rules_dict[str(code)] = dict()
    for idx in specific_rules_i.index:
        specific_rules_dict[str(code)][idx] = specific_rules_i.loc[idx]

etfs_hedge_date_dict = copy.deepcopy(general_rules_dict)
for code, specific_rule in specific_rules_dict.items():
    for mkt, rule in specific_rule.items():
        etfs_hedge_date_dict[code][mkt] = rule

