import pandas as pd
from openpyxl import load_workbook

# 1️⃣ Leer el archivo Excel original
archivo = "proyecto_fun.xlsx"
df = pd.read_excel(archivo, header=None)

# 2️⃣ Convertir los datos en un diccionario para trabajar fácilmente
datos = pd.Series(df[1].values, index=df[0]).to_dict()

# 3️⃣ Extraer valores necesarios
ritmo = float(datos["RITMO(KM/H)"])
peso = float(datos["PESO(KG)"])
tiempo = float(datos["DURACION(MIN)"])

# 4️⃣ Determinar el MET según el ritmo (ejemplo)
if ritmo >= 6 and ritmo < 14:
    MET = 6
elif ritmo >= 14:
    MET = 10
else:
    MET = 3

# 5️⃣ Calcular calorías
calorias = MET * peso * tiempo / 60

print(f"✅ Calorías calculadas: {calorias}")

# 6️⃣ Abrir el archivo Excel para ESCRIBIR el resultado
libro = load_workbook(archivo)
hoja = libro.active

# En tu captura, "CALORIAS QUEMADAS" está en la fila 9 columna A,
# así que el resultado va en la celda B9:
hoja["B9"] = calorias

# 7️⃣ Guardar cambios
libro.save(archivo)
print("📁 Archivo actualizado con el resultado en la celda B9 ✅")
