import xlwings as xlw
import pandas as pd
wb = xlw.Book()
sheet = wb.sheets['Sheet1']
path1 = r'D:\data\customers\IPT\AMGEN\USTO\Lab-Data.xlsx'
sheet1 = 'Sheet1'
df = pd.read_excel(path1, sheet_name=sheet1, index_col = False)
sheet['A1'].options(index=False).value = df
#sheet['A1'].value = ['Sheet Names','Existing IP Address','New IP Address']
#sheet['A2'].value = ['Sheet1','10.116.130.29','10.116.130.6']
#wb = xlw.Book(r'D:\data\customers\IPT\AMGEN\USTO\xlw-test.xlsx')