import pandas as pd
etfs_lst = ['513010', '513180', '513380', '159920', '159866', '513580']
etfs_hedge_date_dict = {
    '513010': {
        'Fixing': '15:40-16:00__T+1',  # close_T+1
        'FX': 'Mostly_Connect_20:25-20:35__T+0'  # Mostly Connect T0
    },
    '513180': {
        'Fixing': '15:40-16:00__T+1',
        'FX': '31%_QDII_15:40-16:00__T+1,69%_Connect_20:25-20:35__T+0'
    },
    '513380': {
        'Fixing': '15:45-15:55__T+0',
        'FX': 'QDIImid_15:40-16:00__T+1'
    },
    '159920': {
        'Fixing': '15:45-15:55__T+1',
        'FX': 'Mostly_Connect_20:25-20:35__T+0'
    },
    '159866': {
        'Fixing': '13:40-14:00__T+1',  # T+1 2pm
        'FX': 'Live_FX_13:40-14:00__T+1'  # Live FX at T+1 close: 
    },
    '513580': {
        'Fixing': '15:40-16:00__T+1',
        'FX': 'QDIImid_15:40-16:00__T+1'
    },
}
etfs_calendar_dict = {
    '513010': ['SSE', 'HKEX'],
    '513180': ['SSE', 'HKEX'],
    '513380': ['SSE', 'HKEX'],
    '159920': ['SSE', 'HKEX'],
    '159866': ['SSE', 'XTKS'],
    '513580': ['SSE', 'HKEX'],
}
start_date = pd.Timestamp('2023-05-01')
end_date = pd.Timestamp('2028-12-31')
# excel_path = 'sheet_automation.xlsx'
excel_path = r"D:\GitHub\spreadsheet_automation_v2\spreadsheet_automation.xlsx"