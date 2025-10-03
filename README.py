import pandas as pd
from openpyxl import load_workbook


df=pd.read_excel("proyecto_fun.xlsx",header=None)
datos=pd.Series(df[1].values,index=df[0]).to_dict()


ritmo=float(datos["RITMO(KM/H)"])
peso=float(datos["PESO(KG)"])
tiempo=float(datos["DURACION(MIN)"])


if ritmo >= 6 and ritmo < 14:
    MET=6
elif ritmo >= 14 and ritmo < 17:
    MET=12.5
elif ritmo >= 17 and ritmo < 20:
    MET=18
else:
    print("no es humano")
    MET=0


calorias =(MET*peso*tiempo)/60
print(f"las calorias quemadas son {calorias}")

wb= load_workbook("proyecto_fun.xlsx")
ws = wb.active

for row in ws.iter_rows(min_row=1,max_col=2):
    if row[0].value and "CALORIAS QUEMADAS" in str(row[0].value).upper():
        row[1].value = calorias
        break
wb.save("proyecto_fun_re.xlsx")
