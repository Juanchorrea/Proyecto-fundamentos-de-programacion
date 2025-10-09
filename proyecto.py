import pandas as pd
from openpyxl import load_workbook
def aña():
    Name=input("ingrese su nombre: ")
    Age=int(input("ingrese su edad: "))
    Peso=int(input("ingrese su peso en kg: "))
    Time=float(input("ingrese el tiempo de ejercicio en minutos: "))
    distance=float(input("ingrese la distancia recorrida en km: "))
    rythm=(distance/(Time/60))
    print(rythm)

    try:
        df = pd.read_excel('proyecto_fun.xlsx')
    except:
        df= pd.DataFrame(columns=["NOMBRE","EDAD","PESO(KG)","DURACION(MIN)","DISTANCIA RECORRIDA(KM)","RITMO(KM/H)"])
    
    new_row = {
    "NOMBRE": Name,
    "EDAD": Age,
    "PESO(KG)": Peso,
    "DURACION(MIN)": Time,
    "DISTANCIA RECORRIDA(KM)": distance,
    "RITMO(KM/H)": rythm
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel('proyecto_fun_re.xlsx', index=False)



    if rythm <= 6 and rythm < 14:
        MET=6
    elif rythm >= 14 and rythm < 17:
        MET=12.5
    elif rythm >= 17 and rythm < 20:
        MET=18
    else:
        print("no es humano")
        MET=0


    calorias =(MET*Peso*Time)/60
    print(f"las calorias quemadas son {calorias:.2f}")

    df["CALORIAS QUEMADAS (KCAL)"] = ""
    df.loc[df.index[-1], "CALORIAS QUEMADAS (KCAL)"] = calorias
    df.to_excel("proyecto_fun_re.xlsx", index=False)
aña()
