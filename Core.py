import json
from datetime import datetime
import requests
from zipfile import ZipFile
from io import BytesIO
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

# daily: D, hourly: H, minutely: T, secondly: S
# see https://pandas.pydata.org/docs/user_guide/timeseries.html#timeseries-offset-aliases
with open("setup.json", encoding="utf-8") as f:
    cfg = json.load(f)
delta = cfg["delta"]
source = cfg["source"]
time = datetime.now().strftime("%Y%m%d%H%M")

L = list()
for i in source:
    url = f'https://join.gov.tw/idea/files/zip/{i["id"]}/export_{time}.zip'
    r = requests.get(url)
    with ZipFile(BytesIO(r.content)) as zf:
        with zf.open("附議名單.csv") as f:
            data = pd.read_csv(f)
            data["附議時間"] = pd.to_datetime(data["附議時間"])
            S = data["附議時間"].dt.floor(freq=delta)
            C = pd.Series(0, index=pd.date_range(S.iloc[0], S.iloc[-1], freq=delta))
            C = C.add(S.value_counts(), fill_value=0).astype(int)
            L.append((data, C))
with pd.ExcelWriter("result.xlsx", engine="xlsxwriter", datetime_format='yyyy-mm-dd hh:mm:ss') as f:
    wb = f.book
    for n, df in enumerate(L, 1):
        sheetname = f"Sheet{n}"
        df[0].to_excel(f, sheetname, startcol=0, startrow=1, header=False, index=False)
        df[1].to_excel(f, sheetname, startcol=7, startrow=1, header=False)
        ws = f.sheets[sheetname]
        for i in range(len(df[1])):
            ws.write(i + 1, 9, f"=SUM({xl_rowcol_to_cell(1, 8)}:{xl_rowcol_to_cell(i+1, 8)})")
        ws.add_table(0, 0, len(df[0]), 5, {'columns': [{'header': column} for column in df[0].columns]})
        ws.add_table(0, 7, len(df[1]), 9, {'columns': [{'header': column} for column in ["時間", "計數", "總數"]]})
        chart = wb.add_chart({'type': 'column'})
        chart.add_series({
            'name': f'={sheetname}!$I$1',
            'categories': f'={sheetname}!$H$2:{xl_rowcol_to_cell(i+1, 7)}',
            'values': f'={sheetname}!$I$2:{xl_rowcol_to_cell(i+1, 8)}'
        })
        chart2 = wb.add_chart({'type': 'scatter', 'subtype': 'straight'})
        chart2.add_series({
            'name': f'={sheetname}!$J$1',
            'categories': f'={sheetname}!$H$2:{xl_rowcol_to_cell(i+1, 7)}',
            'values': f'={sheetname}!$J$2:{xl_rowcol_to_cell(i+1, 9)}',
            'y2_axis': 1
        })
        chart.combine(chart2)
        chart.set_x_axis({'name': '日期', 'num_font': {'rotation': 90}, 'major_gridlines': {'visible': True}})
        chart.set_y_axis({'name': '計數'})
        chart2.set_y2_axis({'name': '總數'})
        chart.set_size({'width': 1080, 'height': 607.5})
        chart.set_legend({'position': 'bottom'})
        try:
            chart.set_title({'name': source[n - 1]["title"]})
        except IndexError:
            pass
        ws.insert_chart('L1', chart)
