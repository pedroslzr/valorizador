# Importamos Pandas como pd
import pandas as pd
import numpy as np
import math
from xlsxwriter.utility import xl_rowcol_to_cell
# from xlsxwriter.utility import write_string

# Herramienta de redondeo a 2 decimales
def round_2dec(n):
    ''' Herramienta de redondeo a 2 decimales
    '''
    multiplier = 10 ** 2
    return math.floor(n*multiplier + 0.5) / multiplier

# Importo datos de dataframe, salto las primeras filas de membrete
df = pd.read_csv("./data/data.csv", skiprows=15, sep=";", encoding='utf-8-sig', index_col=False)
df = df.fillna(0)

# Importo datos del membrete
df2 = pd.read_csv("./data/data.csv", sep=";", encoding='utf-8-sig', index_col=False, header=None)
df2 = df2.iloc[0:14,0:2]
df2 = df2.fillna(0)

# Crear columnas

# Creo columna de presupuesto oferta
df["ppto"] = df["METRADO CONTRACTUAL"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de presupuesto
df["ppto"] = df.apply(lambda x: round_2dec(x["ppto"]), axis=1)

# Costo acumulado anterior
df["costant"] = df["ACUMULADO ANTERIOR"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de costo acumulado anterior
df["costant"] = df.apply(lambda x: round_2dec(x["costant"]), axis=1)

# Valorizacion actual
df["costval"] = df["AVANCE ACTUAL"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de costo actual de valorizacion
df["costval"] = df.apply(lambda x: round_2dec(x["costval"]), axis=1)

# Metrado acumulado actual
df["macum"] = df["ACUMULADO ANTERIOR"] + df["AVANCE ACTUAL"]
# No es necesario aplicar redondeo, son los datos de entrada

# Costo acumulado actual 
df["costacum"] = df["macum"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de costo acumulado
df["costacum"] = df.apply(lambda x: round_2dec(x["costacum"]), axis=1)

# porcentaje % de avance de costo acumulado actual
df["poracum"] = df["costacum"] / df["ppto"]
# truncar a mostrar 2 decimales
df["poracum"] = df["poracum"].apply(lambda x: float("{:.2f}".format(x)))

# Metrado saldo por valorizar
df["msald"] = df["METRADO CONTRACTUAL"] - df["macum"]
# Aplicar el redondeo a la columna de metrado faltante (saldo por valorizar) no es necesario
#df["msald"] = df.apply(lambda x: round_2dec(x["msald"]), axis=1)

# Costo faltante por valorizar (saldo)
df["costsald"] = df["ppto"] - df["costacum"]
#df["costsald"] = df["msald"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de saldo por valorizar
df["costsald"] = df.apply(lambda x: round_2dec(x["costsald"]), axis=1)

# porcentaje % de faltante de avance de proyecto
df["porsald"] = df["costsald"] / df["ppto"]
# truncar a mostrar 2 decimales
df["porsald"] = df["porsald"].apply(lambda x: float("{:.2f}".format(x)))

# Reordenar columnas, definir orden y aplicar
orden = ['ITEM.','DESCRIPCION DE PARTIDA','UND.','METRADO CONTRACTUAL','P.U. OFERTA S/.','ppto','ACUMULADO ANTERIOR',
         'costant','AVANCE ACTUAL','costval','macum','costacum','poracum','msald','costsald','porsald']
df = df[orden]

# Sumas
presupuesto = df["ppto"].sum()
coant = df["costant"].sum()
coval = df["costval"].sum()
coacum = round_2dec(df["costacum"].sum())
pracum = '{:.2%}'.format(coacum / presupuesto) #formatea a % con 2 puntos decimales
cosald = round_2dec(df["costsald"].sum())
prsald = '{:.2%}'.format(cosald / presupuesto) #formatea a % con 2 puntos decimales

#lista_sumas = [presupuesto,coant,coval,coacum,pracum,cosald,prsald]
lista_sumas = [presupuesto,coant,coval,coacum,cosald]
lista_sumas_por = [pracum,prsald]

lista_col_sumas = [5,7,9,11,14]
lista_col_sumas_pr = [12,15]


# Diccionario
# Cambiamos el nombre de las columnas al dataframe original
df.rename(columns = {"ppto": "PRESUPUESTO OFERTA S/.",
                     "ACUMULADO ANTERIOR":"METRADO","costant":"S/.",
                     "AVANCE ACTUAL":"METRADO","costval":"S/.",
                     "macum":"METRADO","costacum":"S/.", "poracum":"%",
                     "msald":"METRADO","costsald":"S/.", "porsald":"%"}, inplace = True)



# sumas
sum_column = df.sum(axis=0)

print(df)
print (sum_column)
print (presupuesto, coant, coval, coacum, pracum, cosald, prsald)


#Manipulacion del excel

df.to_csv('out/out.csv', encoding='utf-8-sig', index=False)
df.to_excel('out/out.xlsx', encoding='utf-8-sig', index=False, sheet_name="valorizacion")
df2.to_excel('out/test.xlsx', encoding='utf-8-sig', index=False, sheet_name="test")

# Acceder al excel
writer = pd.ExcelWriter("out/valorizacion.xlsx", engine='xlsxwriter')
df.to_excel(writer, header=True, startrow=10, startcol=0, index=False, sheet_name="valorizacion")

# Crear el entorno
workbook = writer.book
worksheet = writer.sheets["valorizacion"]

# Ancho de columnas
worksheet.set_zoom(70)

lista_columnas_nombres = ["A:A","B:B","C:C","D:D","E:E","F:F","G:G","H:H","I:I","J:J","K:K","L:L","M:M","N:N","O:O","P:P"]
lista_anchos = [14.57,60,8,14,13,18,14,16,14,16,14,16,11,14,16,11]
for idx2,element3 in enumerate(lista_anchos):
    worksheet.set_column(lista_columnas_nombres[idx2], element3)

# Ancho de filas
worksheet.set_row(1, 40)

# Formatos
titulo_format = workbook.add_format({
    'bold': 0,
    'text_wrap': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    })

subtitulo_format = workbook.add_format({
    'bold': 1,
    'text_wrap': True,
    'border': 0,
    'align': 'center',
    'valign': 'vcenter',
    })

item_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#C4D79B',
    'border': 1})

total_fmt = workbook.add_format({'align': 'right', 'num_format': '$#,##0','bold': True, 'bottom':6})
subtotal_fmt = workbook.add_format({'align': 'right', 'num_format': '$#,##0','bold': False})

por_fmt = workbook.add_format({'align': 'right','bold': True})
porc_fmt = workbook.add_format({'align': 'right','bold': False,'num_format': '#,##0.00"%"'})

currency_format = workbook.add_format({'align': 'right','bold': True,'num_format': '"S/ "#,##0.00'})
currency_format_sub = workbook.add_format({'align': 'right','bold': False,'num_format': '"S/ "#,##0.00'})

number_format = workbook.add_format({'num_format': '#,##0.00'})

# titulos
title = str(df2.iloc[0][1])
numero_val = str(df2.iloc[13][1])
mes_val = str(df2.iloc[11][1])
ano_val = str(df2.iloc[12][1])
subtitulo = "VALORIZACION DE OBRA Nº:  "+numero_val+" CORRESPONDIENTE AL MES DE "+mes_val+ano_val
worksheet.merge_range('A2:P2', title, titulo_format)
worksheet.merge_range('B7:G7', subtitulo, subtitulo_format)

# Encabezado de tabla
lista_columnas = list(df.columns.values)
lista_celdas = ["A9:A11","B9:B11","C9:C11","D9:D11","E9:E11","F9:F11","G10:G11","H10:H11","I10:I11","J10:J11","K10:K11","L10:L11","M10:M11","N10:N11","O10:O11","P10:P11"]
for idx,element in enumerate(lista_columnas):
    worksheet.merge_range(lista_celdas[idx],element,item_format)
  
lista_columnas_merge = ["G9:H9","I9:J9","K9:M9","N9:P9"]
lista_subtitulos = ["ACUMULADO ANTERIOR","AVANCE FISICO","ACUMULADO ACTUAL","SALDO POR VALORIZAR"]
for idx3,element4 in enumerate(lista_subtitulos):
    worksheet.merge_range(lista_columnas_merge[idx3], element4,item_format)

# Sumas en el pie de pagina
number_rows = df.shape[0] + 10

for idx4,element5 in enumerate(lista_sumas):
    worksheet.write_number(number_rows+1,lista_col_sumas[idx4], float(element5),currency_format)

for idx5,element6 in enumerate(lista_sumas_por):
    worksheet.write_string(number_rows+1,lista_col_sumas_pr[idx5], str(element6),por_fmt)

# Resultados pie de pagina
gg = 0.0981015582855174
ut = 0.08
igv = 0.18

lista_gg =list(map(round_2dec,[item*gg for item in lista_sumas]))
for idx6,element7 in enumerate(lista_gg):
    worksheet.write_number(number_rows+2,lista_col_sumas[idx6], float(element7),currency_format_sub)

lista_ut =list(map(round_2dec,[item*ut for item in lista_sumas]))
for idx7,element8 in enumerate(lista_ut):
    worksheet.write_number(number_rows+3,lista_col_sumas[idx7], float(element8),currency_format_sub)

lt_val_sin = [lista_sumas, lista_gg, lista_ut]
lista_val_sin = sum(map(np.array, lt_val_sin))
for idx8,element9 in enumerate(lista_val_sin):
    worksheet.write_number(number_rows+4,lista_col_sumas[idx8], float(element9),currency_format)

lista_igv =list(map(round_2dec,[item*igv for item in lista_val_sin]))
for idx9,element10 in enumerate(lista_igv):
    worksheet.write_number(number_rows+5,lista_col_sumas[idx9], float(element10),currency_format_sub)

lt_val = [lista_val_sin, lista_igv]
lista_val = sum(map(np.array, lt_val))
for idx10,element11 in enumerate(lista_val):
    worksheet.write_number(number_rows+6,lista_col_sumas[idx10], float(element11),currency_format)


# Indice sumas
worksheet.write_string(number_rows+1, 1, "TOTAL COSTO DIRECTO",total_fmt)
worksheet.write_string(number_rows+2, 1, "GASTOS GENERALES",subtotal_fmt)
worksheet.write_number(number_rows+2, 2, float(gg*100),porc_fmt)
worksheet.write_string(number_rows+3, 1, "UTILIDAD",subtotal_fmt)
worksheet.write_number(number_rows+3, 2, float(ut*100),porc_fmt)
worksheet.write_string(number_rows+4, 1, "VALORIZACION MENSUAL (Sin I.G.V)",total_fmt)
worksheet.write_string(number_rows+5, 1, "I.G.V",subtotal_fmt)
worksheet.write_number(number_rows+5, 2, float(igv*100),porc_fmt)
worksheet.write_string(number_rows+6, 1, "TOTAL A VALORIZAR EN EL MES (Incl. I.G.V)",total_fmt)

writer.save()

#TODO
#Interpretar mas datos del membrete
#crear base de datos que acumule datos
#contrastar los avances con la programacion planificada