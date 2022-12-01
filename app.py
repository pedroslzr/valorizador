# Importamos Pandas como pd
import pandas as pd
import numpy as np
import math
import xlsxwriter

# Herramienta de redondeo a 2 decimales
def round_2dec(n):
    ''' Herramienta de redondeo a 2 decimales
    '''
    multiplier = 10 ** 2
    return math.floor(n*multiplier + 0.5) / multiplier

# Importo datos, salto las primeras filas de membrete
df = pd.read_csv("./data/data.csv", skiprows=14, sep=";", encoding='utf-8-sig', index_col=False)
df = df.fillna(0)

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
orden = ['ITEM.','DESCRIPCION DE PARTIDA','UND.','METRADO CONTRACTUAL','P.U. OFERTA S/.','ppto','ACUMULADO ANTERIOR','costant','AVANCE ACTUAL','costval','macum','costacum','poracum','msald','costsald','porsald']
df = df[orden]

# Sumas
presupuesto = df["ppto"].sum()
coant = df["costant"].sum()
coval = df["costval"].sum()
coacum = round_2dec(df["costacum"].sum())
pracum = '{:.2%}'.format(coacum / presupuesto) #formatea a % con 2 puntos decimales
cosald = round_2dec(df["costsald"].sum())
prsald = '{:.2%}'.format(cosald / presupuesto) #formatea a % con 2 puntos decimales


# Diccionario
# Cambiamos el nombre de las columnas al dataframe original
df.rename(columns = {"ppto": "PRESUPUESTO OFERTA S/."}, inplace = True)



#header_dict = {"ppto" : "PRESUPUESTO_CONTRACTUAL"}





sum_column = df.sum(axis=0)

print(df)
print (sum_column)
print (presupuesto, coant, coval, coacum, pracum, cosald, prsald)


#Manipulacion del excel

#borrar columnas
#df.drop(columns = 'label_first_column', axis = 1, inplace= True)
#df.drop(columns = df.columns[0], axis = 1, inplace= True)
# df = [['col4', 'pancho', 'dfgdfg']]

df.to_csv('out/out.csv', encoding='utf-8-sig', index=False)
df.to_excel('out/out.xlsx', encoding='utf-8-sig', index=False, sheet_name="valorizacion")

# Acceder al excel
writer = pd.ExcelWriter("out/valorizacion.xlsx", engine="xlsxwriter")
df.to_excel(writer, index=False, sheet_name="valorizacion")
#
workbook = writer.book
worksheet1 = writer.sheets["valorizacion"]
#
worksheet1.set_zoom(90)

#header1 = '&CHere is some centered text.'
#worksheet1.set_header(header1)



writer.save()