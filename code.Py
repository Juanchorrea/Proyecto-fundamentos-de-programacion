from flask import Flask, render_template, request
import pandas as pd
import os # Importar 'os' para manejar rutas de archivos

# Configuración de la aplicación Flask
app = Flask(__name__, static_folder='static') 

## 1. Ruta de Inicio (GET /)
@app.route('/')
def index():
    # Inicializa TODAS las variables para que la plantilla no falle
    return render_template('front.html', 
                        nombre="", edad="", peso="", tiempo="", distancia="", ritmo="", calorias="") 

## 2. Ruta para Guardar y Calcular (POST /guardar)
@app.route('/guardar', methods=['POST'])
def guardar():
    #  Captura de datos robusta
    name = request.form.get('nombre', '')
    
    # Intenta convertir los datos a float/int, usa 0 o 0.0 si falla
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

    # 3. Lógica de Cálculo
    if tiempo > 0 and distancia > 0:
        ritmo = distancia / (tiempo / 60)
        # Redondear el ritmo para mostrar mejor en el formulario
        ritmo = round(ritmo, 2) 
    else:
        ritmo = 0.0

    # Lógica MET
    if ritmo > 0 and ritmo < 14:
        MET = 6
    elif ritmo >= 14 and ritmo < 17:
        MET = 12.5
    elif ritmo >= 17 and ritmo < 20:
        MET = 18
    elif ritmo >= 20:
        MET = 22 # Valor alto para ritmos extremos
    else:
        MET = 0

    calorias = (MET * peso * tiempo) / 60
    # Redondear las calorías para mostrar mejor en el formulario
    calorias = round(calorias, 2)

    # 4. Lógica de Pandas para guardar en Excel
    filename = 'proyecto_fun_re.xlsx'
    
    try:
        df = pd.read_excel(filename)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            "NOMBRE", "EDAD", "PESO(KG)", "DURACION(MIN)", "DISTANCIA(KM)",
            "RITMO(KM/H)", "CALORIAS(KCAL)"
        ])
    except Exception:
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

    # 5. Renderizar la plantilla, manteniendo INPUTS y mostrando OUTPUTS
    return render_template('front.html', 
                            nombre=name, edad=age, peso=peso, tiempo=tiempo, distancia=distancia,
                            ritmo=ritmo, calorias=calorias)

if __name__ == '__main__':

    app.run(debug=True)
