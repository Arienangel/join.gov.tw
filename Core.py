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
    cfg=json.load(f)
delta = cfg["delta"]
source = cfg["source"]
time = datetime.now().strftime("%Y%m%d%H%M")

L = list()
for i in source:
    url = f'https://join.gov.tw/idea/files/zip/{i["id"]}/export_{time}.zip'
    r = requests.get(url)
    with ZipFile(BytesIO(r.content)) as zf:
        with zf.open("附議名單.csv") as f:
            df = pd.read_csv(f)
            S = pd.to_datetime(df["附議時間"]).dt.floor(freq=delta)
            C = pd.Series(0, index=pd.date_range(S.iloc[0], S.iloc[-1], freq=delta)).add(S.value_counts(),
                                                                                         fill_value=0).astype(int)
            L.append(C)
with pd.ExcelWriter("result.xlsx", engine="xlsxwriter") as f:
    wb = f.book
    for n, df in enumerate(L, 1):
        sheetname = f"Sheet{n}"
        df.to_excel(f, sheetname)
        ws = f.sheets[sheetname]
        ws.write(0, 1, "計數")
        ws.write(0, 2, "總數")
        for i in range(len(df)):
            ws.write(i + 1, 2, f"=SUM({xl_rowcol_to_cell(1, 1)}:{xl_rowcol_to_cell(i+1, 1)})")
        ws.set_column(0, 0, 20.5)
        chart = wb.add_chart({'type': 'column'})
        chart.add_series({
            'name': f'={sheetname}!$B$1',
            'categories': f'={sheetname}!$A$2:{xl_rowcol_to_cell(i+1, 0)}',
            'values': f'={sheetname}!$B$2:{xl_rowcol_to_cell(i+1, 1)}'
        })
        chart2 = wb.add_chart({'type': 'scatter', 'subtype': 'straight'})
        chart2.add_series({
            'name': f'={sheetname}!$C$1',
            'categories': f'={sheetname}!$A$2:{xl_rowcol_to_cell(i+1, 0)}',
            'values': f'={sheetname}!$C$2:{xl_rowcol_to_cell(i+1, 2)}',
            'y2_axis': 1
        })
        chart.combine(chart2)
        chart.set_x_axis({'name': '日期', 'num_font': {'rotation': 90}, 'major_gridlines': {'visible': True}})
        chart.set_y_axis({'name': '計數'})
        chart2.set_y2_axis({'name': '總數'})
        chart.set_size({'width': 1080, 'height': 607.5})
        chart.set_legend({'position': 'bottom'})
        try:
            chart.set_title({'name': source[n-1]["title"]})
        except:
            pass
        ws.insert_chart('D1', chart)
