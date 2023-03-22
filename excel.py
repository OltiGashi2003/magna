import pandas as pd
import openpyxl
import xlwings as xw
import xlsxwriter


mainfile = 'C:/Users/oltigash/PycharmProjects/python_hf/MainFile.xlsx'
subfile = 'C:/Users/oltigash/PycharmProjects/python_hf/Subfile.xlsx'
sheet_name = 'Sheet1' # Name von sheet wo die Daten schreiben sollen
start_row = 8 # von welche zeile soll pandas daten in excel file verfassen

df_y = pd.read_excel(subfile, header=None)
df_y_transposed = df_y.T
df_y_transposed.columns = df_y_transposed.iloc[0]
df_y_transposed = df_y_transposed.drop(df_y_transposed.index[0])


df_main = pd.read_excel(mainfile, skiprows=8, header=None)
df_main.columns = df_main.iloc[0]
df_main = df_main.drop(df_main.index[0])

frames = [df_main, df_y_transposed]
result = pd.concat(frames)
print(result)


result.to_excel(mainfile, index=False)

# Save the changes to the Excel file
writer = pd.ExcelWriter(mainfile, engine='xlsxwriter')
result.to_excel(writer, sheet_name= sheet_name , startrow = start_row, index=False)
writer.save()







