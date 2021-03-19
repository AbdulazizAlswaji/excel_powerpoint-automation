import pandas as pd
import glob
import win32com.client
import time


df = pd.DataFrame()
for file in glob.glob('./data/*.csv'):
    df = pd.concat([df, pd.read_csv(file)])


del df['Unnamed: 0']

df = df.rename(columns={
    'Grade': 'grade',
    'Year': 'year',
    'Category': 'category',
    'Number Tested': 'count'
})


df = df[df.grade != 'All Grades']

pd.DataFrame(df.groupby('year')['count'].sum()).reset_index().to_csv('./output/01.csv', index=False)
pd.DataFrame(df.groupby(['year', 'category'])['count'].sum()).reset_index().to_csv('./output/02.csv', index=False)
pd.DataFrame(df.groupby(['grade', 'category'])['count'].sum()).reset_index().to_csv('./output/03.csv', index=False)

office = win32com.client.Dispatch('Excel.Application')
office.Visible = 0

wb = office.Workbooks.Open('C:\\Users\\A-Scripts\\Desktop\\Report_automation\\main.xlsx')

wb.RefreshAll()
time.sleep(10)

wb.RefreshAll()
time.sleep(10)

count = wb.Sheets.Count
for i in range(count):
    ws = wb.Worksheets[i]
    
    try:
        pivotCount = ws.PivotTables().Count
        for j in range(1, pivotCount + 1):
            ws.PivotTables(j).PivotCache.Refresh()
    except:
        pass
        
        
        
        
wb.Save()
office.Quit()
