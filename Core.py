import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

# daily: D, hourly: H, minutely: T, secondly: S
# see https://pandas.pydata.org/docs/user_guide/timeseries.html#timeseries-offset-aliases
filename = r"附議名單.csv"
delta = "D"

df = pd.read_csv(filename)
S = pd.to_datetime(df["附議時間"]).dt.floor(freq=delta)
C = pd.Series(0, index=pd.date_range(S.iloc[0], S.iloc[-1], freq=delta)).add(S.value_counts(), fill_value=0).astype(int)
with pd.ExcelWriter("result.xlsx", engine="xlsxwriter") as f:
    C.to_excel(f)
    wb = f.book
    ws = wb.worksheets()[0]
    ws.write(0, 1, "計數")
    ws.write(0, 2, "總數")
    for i in range(len(C)):
        ws.write(i + 1, 2, f"=SUM({xl_rowcol_to_cell(1, 1)}:{xl_rowcol_to_cell(i+1, 1)})")
    ws.set_column(0, 0, 20.5)
    chart = wb.add_chart({'type': 'column'})
    chart.add_series({
        'name': '=Sheet1!$B$1',
        'categories': f'=Sheet1!$A$2:{xl_rowcol_to_cell(i+1, 0)}',
        'values': f'=Sheet1!$B$2:{xl_rowcol_to_cell(i+1, 1)}'
    })
    chart2 = wb.add_chart({'type': 'scatter', 'subtype': 'straight'})
    chart2.add_series({
        'name': '=Sheet1!$C$1',
        'categories': f'=Sheet1!$A$2:{xl_rowcol_to_cell(i+1, 0)}',
        'values': f'=Sheet1!$C$2:{xl_rowcol_to_cell(i+1, 2)}',
        'y2_axis': 1
    })
    chart.combine(chart2)
    chart.set_x_axis({'name': '日期', 'num_font': {'rotation': 90}, 'major_gridlines': {'visible': True}})
    chart.set_y_axis({'name': '計數'})
    chart2.set_y2_axis({'name': '總數'})
    chart.set_size({'width': 1080, 'height': 607.5})
    chart.set_legend({'position': 'bottom'})
    ws.insert_chart('D1', chart)
