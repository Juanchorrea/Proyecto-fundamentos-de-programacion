from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__, static_folder='static')

# Ruta de inicio
@app.route('/')
def index():
    return render_template(
        'front.html',
        nombre="", edad="", peso="", tiempo="", distancia="", ritmo="", calorias=""
    )

# Ruta para guardar datos y calcular
@app.route('/guardar', methods=['POST'])
def guardar():
    name = request.form.get('nombre', '')
    try:
        age = int(request.form.get('edad', 0))
    except ValueError:
        age = 0
    try:
        peso = float(request.form.get('peso', 0.0))
    except ValueError:
        peso = 0.0
    try:
        tiempo = float(request.form.get('tiempo', 0.0))
    except ValueError:
        tiempo = 0.0
    try:
        distancia = float(request.form.get('distancia', 0.0))
    except ValueError:
        distancia = 0.0

    # Cálculo del ritmo
    ritmo = distancia / (tiempo / 60) if tiempo > 0 else 0
    ritmo = round(ritmo, 2)

    # Cálculo del MET
    if ritmo > 0 and ritmo < 14:
        MET = 6
    elif ritmo >= 14 and ritmo < 17:
        MET = 12.5
    elif ritmo >= 17 and ritmo < 20:
        MET = 18
    elif ritmo >= 20:
        MET = 22
    else:
        MET = 0

    # Cálculo de calorías
    calorias = (MET * peso * tiempo) / 60
    calorias = round(calorias, 2)

    # Guardar en Excel
    filename = 'proyecto_fun_re.xlsx'
    try:
        df = pd.read_excel(filename)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            "NOMBRE", "EDAD", "PESO(KG)", "DURACION(MIN)", "DISTANCIA(KM)",
            "RITMO(KM/H)", "CALORIAS(KCAL)"
        ])

    nueva_fila = {
        "NOMBRE": name,
        "EDAD": age,
        "PESO(KG)": peso,
        "DURACION(MIN)": tiempo,
        "DISTANCIA(KM)": distancia,
        "RITMO(KM/H)": ritmo,
        "CALORIAS(KCAL)": calorias
    }

    df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
    df.to_excel(filename, index=False)

    return render_template(
        'front.html',
        nombre=name, edad=age, peso=peso, tiempo=tiempo, distancia=distancia,
        ritmo=ritmo, calorias=calorias
    )

if __name__ == '__main__':
    app.run(debug=True)
