import pandas as pd
import glob
import openpyxl
from tkinter import filedialog
import csv

#Usuario seleccionara folder con los archivos
folder_selected=filedialog.askdirectory()
#si no hay folder seleccionado se termina el script
if not folder_selected:
    exit()
    
file_list=glob.glob(folder_selected + '/*.xlsx')
#lista para almacener los archivos
excel_list = []

for file in file_list:
    excel_list.append(pd.read_excel(file, skiprows=2))
    
excel_merged = pd.DataFrame()

for excel_file in excel_list:
    excel_merged=pd.concat([excel_merged, pd.DataFrame.from_records(excel_file)], ignore_index=False)
#excel_merged=pd.concat(excel_list, ignore_index=False)

print(excel_merged.shape)
print(excel_merged.columns)
excel_merged=excel_merged.loc[:,['Title', 'Wk', 'Thu\nAdm 03-Jan', 'Weekend\nAdm', 'Week\nAdm']]

#excel_merged.groupby('Title')

#print(data2.head(10))
#excel_merged.groupby('Title')

excel_merged.to_excel('consolidated_file.xlsx')

# Calcular la suma de 'Thu\nAdm 03-Jan' para el día inicial
initial_day = excel_merged[excel_merged['Wk'] == 1].groupby('Title')['Thu\nAdm 03-Jan'].sum().reset_index()

# Calcular la suma de 'Weekend\nAdm' para el primer fin de semana
first_weekend = excel_merged[excel_merged['Wk'] == 1].groupby('Title')['Weekend\nAdm'].sum().reset_index()

# Calcular la suma de 'Week\nAdm' para la primera semana
first_week = excel_merged[excel_merged['Wk'] == 1].groupby('Title')['Week\nAdm'].sum().reset_index()
#first_week = excel_merged.groupby('Title')['Week\nAdm'].sum().reset_index()

# Calcular la suma de 'Week\nAdm' para la segunda semana
second_week = excel_merged[excel_merged['Wk'] == 2].groupby('Title')['Week\nAdm'].sum().reset_index()

# Fusionar todos los resultados en un solo DataFrame
result = initial_day.merge(first_weekend, on='Title', how='outer').merge(first_week, on='Title', how='outer').merge(second_week, on='Title', how='outer')
result.columns = ['Title', 'Dia inicial', 'Primer Fin de semana', 'Primera semana', 'Segunda semana']

print(result)
result.to_excel('res.xlsx')
#result.to_csv('re.txt', sep=" ", quoting=csv.QUOTE_NONE, escapechar=" ")