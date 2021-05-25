# -*- coding: utf-8 -*-
"""
Created on Sat Jan 16 05:04:47 2021

@author: rclaro
"""

import os
# import time
import pandas as pd
# import numpy as np
import datetime
# import numpy as np
# from openpyxl import load_workbook
# from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
# from openpyxl.utils import get_column_letter
 

# datetime.datetime.strptime
THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(THIS_FOLDER, 'QUAVER.xlsx')
# print(path)
sheets = ['BBDD', 'CURSO', 'DETALLECURSOS']
#dfBD = pd.read_excel(path, sheet_name = sheets[0], skiprows = 0, usecols = 'A:E') 
dfBD = pd.read_excel(path, sheet_name = sheets[0]) 
# pd.read_excel('resultat-elections-2012.xls', sheet_name = 'France entière T1T2', skiprows = 2,  nrows= 5, usecols = 'A:H')
rows = dfBD.shape[0]
# Limpiando Datos del DataFrame Nan, Nat, Etc
dfBD.fillna("", inplace = True)
APELLIDOS_NOMBRES = []
DNI = []
CARGO = []
for index, row in dfBD.iterrows():
  # nombre 
  cadena1 = dfBD.iat[index,1]
  APELLIDOS_NOMBRES.append(cadena1)
  # DNI
  cadena = dfBD.iat[index,2]
  DNI.append(cadena)
  # CARGO
  cadena = dfBD.iat[index,3]
  CARGO.append(cadena)
Cols = []
for col in dfBD.columns: 
    Cols.append(col)
# print(Cols [4])
# print(len(Cols))
column = 4 
i = 1
id_sap_persona=[]
curso_excel=[]
c_curso = []
c_fecha = []
c_id_curso = []
c_asistentes = []
c_descripcion_curso = []
lfc = []
for j in range(4,len(Cols)):
    curso = Cols [j]
    # 'dfBD[curso].replace('', np.nan, inplace=True)
    dfx = dfBD[curso]
    dfx = dfx.to_frame()
    #print(curso)
    #print(j)
    dfx[curso] = pd.to_datetime(dfx[curso])
    dfx = dfx.sort_values(by=curso)
    dfx = dfx.sort_values(by=curso,ascending=True)
    df  = dfx[curso].unique()
    # df  = df.fillna("", inplace = True)
    for x in range(len(df)): 
        dato = str(df[x] )
        if dato != "NaT" :
            cip = "C" + str(i)
            i +=1
            # print (df[x])
            fecha=df[x]
            ts = pd.to_datetime(str(fecha)) 
            d = ts.strftime('%d/%m/%Y')
            #print(fecha)
            #dtpFecha= fecha.strftime("%d/%m/%Y")
            # print(fecha, d)
            c_curso.append(cip)
            c_fecha.append(d)
            c_descripcion_curso.append(curso)
            dfCurso = dfBD[curso] == fecha
            dfCurso = dfCurso.to_frame()
            dfCurso = dfCurso[dfCurso[curso] == True]
            asistentes = dfCurso.shape[0]
            c_asistentes.append(asistentes)
            if curso.find("(LIFE CRITICAL)") == -1:
                lfc.append("NO")
            else:
                lfc.append("SI")
            for index, row in dfCurso.iterrows():
                cadena = dfBD.iat[index,j]
                dfBD.iat[index,j] = cip
                id_sap_persona.append(dfBD.iat[index,2])
                curso_excel.append(cip)
           # ' print(cadena)
dfCurso_Excel = pd.DataFrame(list(zip(c_descripcion_curso, c_fecha, c_curso, c_asistentes, lfc )), 
                                 #         1               2         3       4           5               6    7     8
                        columns =['CURSO','FECHA', 'IDCURSO', 'CANTIDAD', 'LIFE CRITICAL'])  
# CURSO	FECHA	HORAS	CAPACITADOR	IDSAP CAPACITADOR	IDCURSO	CANTIDAD	IDSAP CURSO	CANTIDAD SAP
dfCurso_Excel.insert(2,'ID SAP CAPACITADOR','')
dfCurso_Excel.insert(2,'CAPACITADOR','')
dfCurso_Excel.insert(2,'HORAS','')
dfCurso_Excel.insert(7,'CANTIDAD SAP','')
dfCurso_Excel.insert(7,'IDSAP CURSO','')
# dfCurso_Excel.insert(4,'CANTIDAD SAP','')
dfDetalleC  = pd.DataFrame(list(zip(id_sap_persona, curso_excel)), 
#         1            2         3     4    5    6    7     8
columns =['IDSAP', 'CURSO'])  

dfBBDD  = pd.DataFrame(list(zip(DNI, APELLIDOS_NOMBRES, DNI, DNI, CARGO )), 
#                               1            2           3     4    5    
columns =['N°', 'APELLIDOS Y NOMBRES', 'DNI', 'IDSAP', 'CARGO']) 
#          1            2                 3      4        5    

        # print(type(fecha))
    # if df[x] != NaT:
    #     print("Nat encontrado")
name_report = "BBDDQUAVER.xlsx"
# dfDetalleC.index.names = ['Nº']
dfCurso_Excel.reset_index(drop=True, inplace=True) 
dfCurso_Excel.index = dfCurso_Excel.index + 1
dfCurso_Excel.index.names = ['Nº']
# Create a Pandas Excel writer using XlsxWriter as the engine.
# dfDetalleC.reset_index(drop=True, inplace=False) 
# dfDetalleC.reset_index(drop=True, inplace=True)
writer = pd.ExcelWriter(name_report, engine='xlsxwriter')
# Convert the dataframe to an XlsxWriter Excel object.
# dfDetalleC.dfDetalleC(writer, sheet_name='DETALLECURSOS')
dfCurso_Excel.to_excel(writer, sheet_name='CURSO', index = True)
dfDetalleC.to_excel(writer, sheet_name='DETALLECURSOS', index = False)
dfBBDD.to_excel(writer, sheet_name='BBDD', index = False)
# Close the Pandas Excel writer and output the Excel file.
writer.save()
print("Exportación finalizada")
        
# for index, row in df.iterrows():
#     print(row)
#print(dfBD.sample(10))
# df =  pd.notnull(dfBD['TRABAJOS EN ALTURA'])
# dfa= dfBD[df]
# print("Datos filtrados")
# #                           yyyy-mm-dd 
# start_date = np.datetime64('2019-07-21')
# dfCurso = dfa['TRABAJOS EN ALTURA'] == start_date
# dfCurso = dfCurso.to_frame()
# dfCurso = dfCurso[dfCurso['TRABAJOS EN ALTURA'] == True]
# CIP = "C1"

    
# print(dfCurso.head())
# # df.loc[:, dfBool.values
