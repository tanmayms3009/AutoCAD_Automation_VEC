from itertools import count
import pandas as pd
import numpy as np

df = pd.read_excel('D:\Tanmay\Jobs\Vectre-24.01.22\Automation\Calculate elements\Windows.xlsx', 'Details')

df_table = df[['Elevation', 'Window Type', 'Levels']]
print(df_table)

print("==============================================================")

print(df['Levels'].value_counts('Levels'))

print("==============================================================")

df_pivot = df.pivot_table(index=['Elevation', 'Window Type'], columns='Levels', values='Quantity', aggfunc=[sum], fill_value=0, margins=True, margins_name='Total')
print(df_pivot)

print("==============================================================")

df_windows = df.pivot_table(index=['Window Type'], columns=['Levels'], values='Quantity', aggfunc=[sum], fill_value=0, margins=True, margins_name='Total')
#ADD TOTALS ROW AND COLUMN


print(df_windows)

#df.to_excel(path, sheet_name='...')

with pd.ExcelWriter('Windows_New.xlsx') as writer:
    #GeneralTable
    df_table.to_excel(writer, sheet_name="Details")
    #Elevation wise distribution table
    df_pivot.to_excel(writer, sheet_name="Pivot")
    #Window type wise distribution
    df_windows.to_excel(writer, sheet_name="Window Types")



# pivot_tble=pd.pivot_table(dataframe,index=['Elevation','Window Type'] ,columns=['Levels'])
# print(pivot_tble)
# , aggfunc=[np.sum], fill_value=0
# values=dataframe.groupby('Levels').count()