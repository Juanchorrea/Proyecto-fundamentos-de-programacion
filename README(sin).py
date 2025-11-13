import pandas as pd
from openpyxl import load_workbook

archivo = "proyecto_fun.xlsx"
df = pd.read_excel(archivo, header=None)

datos = pd.Series(df[1].values, index=df[0]).to_dict()

ritmo = float(datos["RITMO(KM/H)"])
peso = float(datos["PESO(KG)"])
tiempo = float(datos["DURACION(MIN)"])

if ritmo >= 6 and ritmo < 14:
    MET = 6
elif ritmo >= 14:
    MET = 10
else:
    MET = 3


calorias = MET * peso * tiempo / 60

print(f" Calorías calculadas: {calorias}")

libro = load_workbook(archivo)
hoja = libro.active

# En tu captura, "CALORIAS QUEMADAS" está en la fila 9 columna A,
# así que el resultado va en la celda B9:
hoja["B9"] = calorias


libro.save(archivo)
print("Archivo actualizado con el resultado en la celda B9 ")
