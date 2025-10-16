# --- Cálculo de ICA (versión extendida con 16 parámetros) ---

def qi_ph(x):
    if 6.5 <= x <= 8.5: return 100
    elif 5.5 <= x < 6.5 or 8.5 < x <= 9.5: return 80
    elif 4.5 <= x < 5.5 or 9.5 < x <= 10.5: return 40
    else: return 20

def qi_temp(x):
    if x <= 25: return 100
    elif x <= 30: return 100 - (x - 25) * 4
    elif x <= 35: return 80 - (x - 30) * 10
    else: return 30

def qi_co2(x):
    if x <= 5: return 100
    elif x <= 10: return 100 - (x - 5) * 10
    elif x <= 20: return 50 - (x - 10) * 5
    else: return 10

def qi_od(x):
    if x >= 6: return 100
    elif x >= 5: return 80
    elif x >= 4: return 60
    elif x >= 2: return 30
    else: return 10

def qi_salinidad(x):
    if x <= 35: return 100
    elif x <= 40: return 90 - (x - 35) * 5
    elif x <= 45: return 65 - (x - 40) * 7
    else: return 30

def qi_alk_fenol(x):
    if x <= 30: return 100
    elif x <= 100: return 100 - (x - 30) * 0.5
    else: return 65 - (x - 100) * 0.3

def qi_alk_total(x):
    if x <= 200: return 100
    elif x <= 400: return 100 - (x - 200) * 0.2
    else: return 60 - (x - 400) * 0.1

def qi_nitrito(x):
    if x <= 0.05: return 100
    elif x <= 0.1: return 100 - (x - 0.05) * 800
    elif x <= 0.5: return 60 - (x - 0.1) * 100
    else: return 20

def qi_nitrato(x):
    if x <= 1: return 100
    elif x <= 5: return 100 - (x - 1) * 20
    elif x <= 10: return 20 - (x - 5) * 4
    else: return 0

def qi_fosfato(x):
    if x <= 0.1: return 100
    elif x <= 0.5: return 100 - (x - 0.1) * 150
    elif x <= 1: return 40 - (x - 0.5) * 40
    else: return 20

def qi_turbidez(x):
    if x <= 1: return 100
    elif x <= 5: return 100 - (x - 1) * 15
    elif x <= 10: return 40 - (x - 5) * 8
    else: return 0

def qi_acidez_nm(x):
    if x <= 20: return 100
    elif x <= 50: return 100 - (x - 20) * 2
    else: return 40

def qi_acidez_fp(x):
    if x <= 10: return 100
    elif x <= 30: return 100 - (x - 10) * 2
    else: return 60

def qi_sst(x):
    if x <= 10: return 100
    elif x <= 25: return 100 - (x - 10) * 2
    elif x <= 50: return 70 - (x - 25) * 2
    else: return 20

def qi_conductividad(x):
    if x <= 15: return 100
    elif x <= 25: return 100 - (x - 15) * 3
    elif x <= 35: return 70 - (x - 25) * 4
    else: return 30


# --- Pesos ---
pesos = {
    'ph': 0.08, 'temperatura': 0.06, 'co2': 0.05, 'od': 0.12, 'salinidad': 0.05,
    'alc_fenolftaleina': 0.04, 'alc_total': 0.04, 'nitrito': 0.05,
    'nitrato': 0.05, 'fosfato': 0.08, 'turbidez': 0.08, 'acidez_nm': 0.04,
    'acidez_fp': 0.04, 'sst': 0.09, 'conductividad': 0.08
}

# --- Importaciones ---
import pandas as pd
import os

def procesar_archivo_excel():
    try:
        # Obtener la ruta a Documents
        ruta_documents = os.path.join(os.path.expanduser('~'), 'Documents')
        ruta_archivo = os.path.join(ruta_documents, 'datos.xlsx')
        
        print(f"\nBuscando archivo en: {ruta_documents}")
        
        # Verificar si el archivo existe
        if os.path.exists(ruta_archivo):
            print("Archivo encontrado!")
            df = pd.read_excel(ruta_archivo)
            print("\nColumnas encontradas:")
            print(df.columns.tolist())
            return df
        else:
            print("\nError: No se encontró el archivo 'datos.xlsx'")
            print("Asegúrate de que:")
            print(f"1. El archivo se llame exactamente 'datos.xlsx'")
            print(f"2. El archivo esté en: {ruta_documents}")
            return None
    except Exception as e:
        print(f"\nError al procesar el archivo: {str(e)}")
        return None

def calcular_ica_muestra(datos):
    # Calcular Qi para cada parámetro
    try:
        Qi = {
            'ph': qi_ph(datos['pH']),
            'temperatura': qi_temp(datos['Temperatura']),
            'co2': qi_co2(datos['CO2']),
            'od': qi_od(datos['OD']),
            'salinidad': qi_salinidad(datos['Salinidad']),
            'alc_fenolftaleina': qi_alk_fenol(datos['Alcalinidad F']),
            'alc_total': qi_alk_total(datos['Alcalinidad T']),
            'nitrito': qi_nitrito(datos['Nitrito']),
            'nitrato': qi_nitrato(datos['Nitrato']),
            'fosfato': qi_fosfato(datos['Fosfato']),
            'turbidez': qi_turbidez(datos['Turbidez']),
            'acidez_nm': qi_acidez_nm(datos['Acidez NM']),
            'acidez_fp': qi_acidez_fp(datos['Acidez FP']),
            'sst': qi_sst(datos['SST']),
            'conductividad': qi_conductividad(datos['Conductividad'])
        }
        
        # Calcular ICA
        ICA = sum(Qi[p] * pesos[p] for p in pesos)
        
        # Determinar categoría
        if ICA > 90: calidad = "Excelente"
        elif ICA > 70: calidad = "Buena"
        elif ICA > 50: calidad = "Regular"
        elif ICA > 25: calidad = "Mala"
        else: calidad = "Muy mala"
        
        return ICA, calidad, Qi
    except KeyError as e:
        print(f"\nError: No se encontró la columna {str(e)}")
        print("Columnas esperadas: pH, Temperatura, CO2, OD, Salinidad, Alcalinidad F, Alcalinidad T,")
        print("Nitrito, Nitrato, Fosfato, Turbidez, Acidez NM, Acidez FP, SST, Conductividad")
        return None, None, None

def main():
    # Leer archivo Excel
    df = procesar_archivo_excel()
    
    if df is None:
        return
        
    print("\nProcesando datos...")
    
    # Procesar cada fila del DataFrame
    resultados = []
    for index, row in df.iterrows():
        print(f"\nProcesando muestra {index + 1}...")
        ica, calidad, qi = calcular_ica_muestra(row)
        
        if ica is not None:
            resultados.append({
                'Muestra': index + 1,
                'ICA': ica,
                'Calidad': calidad,
                'Qi': qi
            })
    
    # Mostrar resultados
    if resultados:
        print("\n=== RESULTADOS DEL ICA POR MUESTRA ===")
        print("\nResumen:")
        print("-" * 50)
        print(f"{'Muestra':^10} {'ICA':^10} {'Calidad':^15}")
        print("-" * 50)
        
        for r in resultados:
            print(f"{r['Muestra']:^10} {r['ICA']:^10.2f} {r['Calidad']:^15}")
        
        print("\nDetalles de subíndices Qi por muestra:")
        for r in resultados:
            print(f"\n--- Muestra {r['Muestra']} ---")
            for param, valor in r['Qi'].items():
                print(f"{param:20}: {valor:5.1f}")
            print(f"ICA total: {r['ICA']:.2f}")
            print(f"Categoría: {r['Calidad']}")

if __name__ == "__main__":
    main()
