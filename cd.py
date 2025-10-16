from flask import Flask, render_template, request, redirect, url_for
import pandas as pd

app = Flask(__name__)

@app.route('/')
def index():
    # Inicializa todas las variables que usa front.html (incluyendo las de input)
    return render_template('front.html', 
                        nombre="", edad="", peso="", tiempo="", distancia="", ritmo="", calorias="")

@app.route('/guardar', methods=['POST'])
def guardar():
    # 1. Usar .get() para evitar KeyError y proporcionar un valor por defecto (ej. None o una cadena vacía)
    # 2. Validar que los campos no estén vacíos antes de intentar la conversión a número
    
    name = request.form.get('nombre', '')
    
    # Intenta obtener el valor, si no existe o está vacío, usa 0 como valor seguro
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

    # Evitar división por cero si el tiempo es 0
    if tiempo > 0:
        ritmo = distancia / (tiempo / 60)
    else:
        ritmo = 0.0

    # Lógica de cálculo (sin cambios)
    if ritmo <= 6 and ritmo < 14:
        MET = 6
    elif ritmo >= 14 and ritmo < 17:
        MET = 12.5
    elif ritmo >= 17 and ritmo < 20:
        MET = 18
    else:
        MET = 0

    calorias = (MET * peso * tiempo) / 60

    # Lógica de DataFrame (sin cambios)
    try:
        df = pd.read_excel('proyecto_fun_re.xlsx')
    except:
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
    df.to_excel('proyecto_fun_re.xlsx', index=False)

    return render_template('front.html', ritmo=ritmo, calorias=calorias)

if __name__ == '__main__':
    app.run(debug=True) # Puedes ejecutar con debug=True para ver errores en detalle