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
#agregar todos los archivos excel a una lista, se elimina las primeras 2 filas de cada archivo de excel
for file in file_list:
    excel_list.append(pd.read_excel(file, skiprows=2))
    
#se crea el dataframe excel_merged
excel_merged = pd.DataFrame()

#consolidando todos los archivos excel en un solo dataframe
excel_merged=pd.concat(excel_list, ignore_index=False)

print(excel_merged.shape)
print(excel_merged.columns)
#Seleccionamos solo las columnas que necesitamos para el procesamiento de datos
excel_merged=excel_merged.loc[:,['Title', 'Wk', 'Thu\nAdm 03-Jan', 'Fri\nAdm 04-Jan', 'Sat\nAdm 05-Jan', 
                         'Sun\nAdm 06-Jan', 'Weekend\nAdm', 'Mon\nAdm 07-Jan', 'Tue\nAdm 08-Jan', 
                         'Wed\nAdm 09-Jan', 'Week\nAdm']]

#Archivo excel consolidado
excel_merged.to_excel('consolidado.xlsx')

# Calcular la suma de 'Thu\nAdm 03-Jan' para el día inicial
initial_day = excel_merged[excel_merged['Wk'] == 1].groupby('Title')['Thu\nAdm 03-Jan'].sum().reset_index()

# Calcular la suma de 'Weekend\nAdm' para el primer fin de semana
first_weekend = excel_merged[excel_merged['Wk'] == 1].groupby('Title')['Weekend\nAdm'].sum().reset_index()

# Calcular la suma de 'Week\nAdm' para la primera semana
first_week = excel_merged[excel_merged['Wk'] == 1].groupby('Title')['Week\nAdm'].sum().reset_index()
#first_week = excel_merged.groupby('Title')['Week\nAdm'].sum().reset_index()

# Calcular la suma de 'Week\nAdm' para la segunda semana
second_week = excel_merged[excel_merged['Wk'] == 2].groupby('Title')['Week\nAdm'].sum().reset_index()

# Calcular la suma de 'Thu\nAdm 03-Jan', 'Fri\nAdm 04-Jan', 'Sat\nAdm 05-Jan', 'Sun\nAdm 06-Jan', 'Weekend\nAdm', 'Mon\nAdm 07-Jan', 'Tue\nAdm 08-Jan', 'Wed\nAdm 09-Jan', 'Week\nAdm' para el total
total = excel_merged.groupby('Title')[['Thu\nAdm 03-Jan', 'Fri\nAdm 04-Jan', 'Sat\nAdm 05-Jan', 'Sun\nAdm 06-Jan', 'Weekend\nAdm', 'Mon\nAdm 07-Jan', 'Tue\nAdm 08-Jan', 'Wed\nAdm 09-Jan', 'Week\nAdm']].sum().reset_index()
total['Total']=total[['Thu\nAdm 03-Jan', 'Fri\nAdm 04-Jan', 'Sat\nAdm 05-Jan', 'Sun\nAdm 06-Jan', 'Weekend\nAdm', 'Mon\nAdm 07-Jan', 'Tue\nAdm 08-Jan', 'Wed\nAdm 09-Jan', 'Week\nAdm']].sum(axis=1)
total_movie=total[['Title','Total']]

"""
Calcular la suma de 'Thu\nAdm 03-Jan', 'Fri\nAdm 04-Jan', 'Sat\nAdm 05-Jan', 'Sun\nAdm 06-Jan', 'Weekend\nAdm', 'Mon\nAdm 07-Jan', 'Tue\nAdm 08-Jan', 'Wed\nAdm 09-Jan', 'Week\nAdm' para el valor restante
remaining = excel_merged[excel_merged['Wk'] >= 3].groupby('Title')[['Thu\nAdm 03-Jan', 'Fri\nAdm 04-Jan', 'Sat\nAdm 05-Jan', 'Sun\nAdm 06-Jan', 'Weekend\nAdm', 'Mon\nAdm 07-Jan', 'Tue\nAdm 08-Jan', 'Wed\nAdm 09-Jan', 'Week\nAdm']].sum().reset_index()
remaining['remaing_total']=remaining[['Thu\nAdm 03-Jan', 'Fri\nAdm 04-Jan', 'Sat\nAdm 05-Jan', 'Sun\nAdm 06-Jan', 'Weekend\nAdm', 'Mon\nAdm 07-Jan', 'Tue\nAdm 08-Jan', 'Wed\nAdm 09-Jan', 'Week\nAdm']].sum(axis=1)
remainig_movie=remaining[['Title','remaing_total']]
print(remainig_movie)
"""

# Fusionar todos los resultados en un solo DataFrame
result = initial_day.merge(first_weekend, on='Title', how='outer').merge(first_week, on='Title', how='outer').merge(second_week, on='Title', how='outer').merge(total_movie, on='Title', how='outer')
result.columns = ['Title', 'Dia inicial', 'Primer Fin de semana', 'Primera semana', 'Segunda semana', 'Total']
result = result.fillna(0)
result['Restante']=result['Total']-(result['Primera semana']+result['Segunda semana'])
result['Primera Semana PCT']=(result['Primera semana']/result['Total']*100).round(2).astype(str) + '%'
result['Segunda semana PCT']=(result['Segunda semana']/result['Total']*100).round(2).astype(str) + '%'
result['Restante PCT']=(result['Restante']/result['Total']*100).round(2).astype(str) + '%'
result=result.reindex(columns=['Title', 'Dia inicial', 'Primer Fin de semana', 'Primera semana', 'Primera Semana PCT', 'Segunda semana', 'Segunda semana PCT', 'Restante','Restante PCT','Total'])
print(result)
result.to_excel('resultado.xlsx')
#result.to_csv('re.txt', sep=" ", quoting=csv.QUOTE_NONE, escapechar=" ")