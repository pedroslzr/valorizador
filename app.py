# Importamos Pandas como pd
import pandas as pd
import numpy as np
import math
import openpyxl

# Herramienta de redondeo a 2 decimales
def round_2dec(n):
    ''' Herramienta de redondeo a 2 decimales
    '''
    multiplier = 10 ** 2
    return math.floor(n*multiplier + 0.5) / multiplier

# Importo datos, salto las primeras filas de membrete
df = pd.read_csv("./data/data.csv", skiprows=14, sep=";", encoding='utf-8-sig')
df = df.fillna(0)

# Crear columnas
# Creo columna de PRESUPUESTO OFERTA  
df["ppto"] = df["METRADO CONTRACTUAL"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de presupuesto


#def apply_cosas(x, col_name: str):
#    return round_2dec(x[col_name])

def apply_roundppto(x):
    return round_2dec(x["ppto"])

df["ppto"] = df.apply(apply_roundppto,axis = 1)

# Costo acumulado anterior
df["costant"] = df["ACUMULADO ANTERIOR"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de costo acumulado anterior
def apply_roundcant(x):
    return round_2dec(x["costant"])
df["costant"] = df.apply(apply_roundcant,axis = 1)

# Valorizacion actual
df["costval"] = df["AVANCE ACTUAL"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de costo actual de valorizacion
def apply_roundcval(x):
    return round_2dec(x["costval"])
df["costval"] = df.apply(apply_roundcval,axis = 1)

# Metrado acumulado actual
df["macum"] = df["ACUMULADO ANTERIOR"] + df["AVANCE ACTUAL"]
# No es necesario aplicar redondeo, son los datos de entrada

# Costo acumulado actual 
df["costacum"] = df["macum"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de costo acumulado
def apply_roundcacum(x):
    return round_2dec(x["costacum"])
df["costacum"] = df.apply(apply_roundcacum,axis = 1)

# porcentaje % de avance de costo acumulado actual
df["poracum"] = df["costacum"] / df["ppto"]
# truncar a 2 decimales
df["poracum"] = df["poracum"].apply(lambda x: float("{:.2f}".format(x)))

# Metrado saldo por valorizar
#df["msald"] = df[]




# Sumas
presupuesto = df["ppto"].sum()
coant = df["costant"].sum()
coval = df["costval"].sum()
coacum = round_2dec(df["costacum"].sum())
pracum = '{:.2%}'.format(coacum / presupuesto) #formatea a % con 2 puntos decimales
# Diccionario
#header_dict = {"ppto" : "PRESUPUESTO_CONTRACTUAL"}
sum_column = df.sum(axis=0)

print(df)
print (sum_column)
print (presupuesto, coant, coval, coacum, pracum)

df.to_csv('out/out.csv', encoding='utf-8-sig')
df.to_excel('out/out.xlsx', encoding='utf-8-sig')

# df = [['col4', 'pancho', 'dfgdfg']]