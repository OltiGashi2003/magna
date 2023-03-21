import pandas as pd
import openpyxl
import xlwings as xw
import xlsxwriter


mainfile = 'MainFile.xlsx'
subfile = 'Subfile.xlsx'

df_y = pd.read_excel(subfile, header=None)
df_y_transposed = df_y.T
df_y_transposed.columns = df_y_transposed.iloc[0]
df_y_transposed = df_y_transposed.drop(df_y_transposed.index[0])

df_main = pd.read_excel(mainfile, header=None)
df_main.columns = df_main.iloc[0]
df_main = df_main.drop(df_main.index[0])

frames = [df_main, df_y_transposed]
result = pd.concat(frames)
result.to_excel(mainfile, index=False)

# Save the changes to the Excel file
writer = pd.ExcelWriter(mainfile, engine='xlsxwriter')
result.to_excel(writer, index=False)
writer.save()







