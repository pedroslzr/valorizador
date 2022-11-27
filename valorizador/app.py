# Importamos Pandas como pd
import pandas as pd
import numpy as np
import math

# Herramienta de redondeo a 2 decimales
def round_2dec(n):
    ''' Herramienta de redondeo a 2 decimales
    '''
    multiplier = 10 ** 2
    return math.floor(n*multiplier + 0.5) / multiplier

# Importo datos, salto las primeras filas de membrete
df = pd.read_csv("./data/data.csv", skiprows=14, sep=";", encoding='utf-8-sig')


# Creo columna de PRESUPUESTO OFERTA
    #df["ppto"] = np.round(df["METRADO CONTRACTUAL"] * df["P.U. OFERTA S/."],2)
df["ppto"] = df["METRADO CONTRACTUAL"] * df["P.U. OFERTA S/."]
    #df["ppto"] = np.round(df["ppto"],2)
# Aplicar el redondeo a la columna de presupuesto
def apply_roundppto(x):
    return round_2dec(x["ppto"])
df["ppto"] = df.apply(apply_roundppto,axis = 1)

# Acumulado anterior> Columna de costo
df["costacum"] = df["ACUMULADO ANTERIOR"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de costo acumulado anterior
def apply_roundcacum(x):
    return round_2dec(x["costacum"])
df["costacum"] = df.apply(apply_roundcacum,axis = 1)

# Acumulado actual> Columna de costo de la valorizacion actual
df["costact"] = df["AVANCE ACTUAL"] * df["P.U. OFERTA S/."]
# Aplicar el redondeo a la columna de costo acumulado anterior
def apply_roundcact(x):
    return round_2dec(x["costact"])
df["costact"] = df.apply(apply_roundcact,axis = 1)

#df["ppto"] = round_half_up(df["ppto"],2)
#df["ppto"] = [round_half_up(n, 2) for n in df["ppto"]]
#df["ppto"] = redondear(df["ppto"])
presupuesto = df["ppto"].sum()
coacum = df["costacum"].sum()
coact = df["costact"].sum()
# Diccionario
#header_dict = {"ppto" : "PRESUPUESTO_CONTRACTUAL"}

print(df)
print(presupuesto)
print(coacum)
print(coact)

df.to_csv('out/out.csv', encoding='utf-8-sig')



# df = [['col4', 'pancho', 'dfgdfg']]