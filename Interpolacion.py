import pandas as pd
import matplotlib.pyplot as plt
import os

class InterpolacionLagrange:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.n = len(x)
    
    def interpolar_punto(self, x_eval, indices_puntos):
        """Interpola un punto usando los índices especificados"""
        resultado = 0.0
        n = len(indices_puntos)
        
        for i in range(n):
            termino = self.y[indices_puntos[i]]
            
            for j in range(n):
                if i != j:
                    denominador = self.x[indices_puntos[i]] - self.x[indices_puntos[j]]
                    if abs(denominador) < 1e-12:
                        return self.y[indices_puntos[i]]
                    termino *= (x_eval - self.x[indices_puntos[j]]) / denominador
            
            resultado += termino
        
        return resultado
    
    def obtener_puntos_para_grado(self, punto_excluir, grado):
        """Obtiene puntos únicos para interpolación"""
        n_puntos_necesarios = grado + 1
        
        if self.n <= n_puntos_necesarios:
            return [i for i in range(self.n) if i != punto_excluir]
        
        puntos_seleccionados = []
        x_vistos = []
        
        for i in range(self.n):
            if i == punto_excluir:
                continue
                
            es_unico = True
            for x_visto in x_vistos:
                if abs(self.x[i] - x_visto) < 1e-12:
                    es_unico = False
                    break
            
            if es_unico:
                puntos_seleccionados.append(i)
                x_vistos.append(self.x[i])
                
            if len(puntos_seleccionados) == n_puntos_necesarios:
                break
        
        if len(puntos_seleccionados) < n_puntos_necesarios:
            puntos_seleccionados = []
            for i in range(min(n_puntos_necesarios, self.n)):
                if i != punto_excluir:
                    puntos_seleccionados.append(i)
        
        return puntos_seleccionados
    
    def calcular_error_grado(self, grado):
        """Calcula error para un grado específico"""
        errores = []
        y_interpolados = []
        
        for i in range(self.n):
            try:
                puntos_interpolacion = self.obtener_puntos_para_grado(i, grado)
                
                if len(puntos_interpolacion) < grado + 1:
                    errores.append(100.0)
                    y_interpolados.append(self.y[i])
                    continue
                
                y_interp = self.interpolar_punto(self.x[i], puntos_interpolacion)
                y_interpolados.append(y_interp)
                
                if abs(self.y[i]) > 1e-12:
                    error = abs((self.y[i] - y_interp) / self.y[i]) * 100
                else:
                    error = 0.0 if abs(y_interp) < 1e-12 else 100.0
                
                errores.append(error)
                
            except Exception:
                errores.append(100.0)
                y_interpolados.append(self.y[i])
        
        return self.x, errores, y_interpolados
    
    def mostrar_resultados(self):
        """Muestra resultados de interpolación"""
        print("\n" + "="*70)
        print("INTERPOLACIÓN DE LAGRANGE")
        print("="*70)
        
        for grado in [1, 2, 3, 4]:
            print(f"\n{'─'*70}")
            print(f"GRADO {grado}")
            print(f"{'─'*70}")
            
            x_puntos, errores, y_interp = self.calcular_error_grado(grado)
            
            print(f"{'x':<10} {'y_real':<12} {'y_interp':<12} {'error %':<12}")
            print("-" * 50)
            
            for i in range(len(x_puntos)):
                print(f"{x_puntos[i]:<10.4f} {self.y[i]:<12.4f} {y_interp[i]:<12.4f} {errores[i]:<12.4f}")
    
    def graficar_errores(self):
        """Genera gráficas de errores para todos los grados"""
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
        
        grados = [1, 2, 3, 4]
        colores = ['red', 'blue', 'green', 'purple']
        marcadores = ['o', 's', '^', 'D']
        
        # Gráfica 1: Errores porcentuales
        for idx, grado in enumerate(grados):
            x_puntos, errores, _ = self.calcular_error_grado(grado)
            
            x_filtrados = []
            errores_filtrados = []
            for i in range(len(errores)):
                if errores[i] <= 100:
                    x_filtrados.append(x_puntos[i])
                    errores_filtrados.append(errores[i])
            
            ax1.plot(x_filtrados, errores_filtrados, 
                    color=colores[idx], 
                    marker=marcadores[idx],
                    linewidth=2,
                    markersize=6,
                    label=f'Grado {grado}',
                    alpha=0.7)
        
        ax1.set_xlabel('Coordenada x')
        ax1.set_ylabel('Error Porcentual (%)')
        ax1.set_title('Comparación de Errores en Interpolación de Lagrange')
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        
        # Gráfica 2: Datos originales vs interpolados
        for idx, grado in enumerate(grados):
            x_puntos, _, y_interp = self.calcular_error_grado(grado)
            
            ax2.plot(x_puntos, y_interp,
                    color=colores[idx],
                    marker=marcadores[idx],
                    linewidth=2,
                    markersize=4,
                    label=f'Interpolación Grado {grado}',
                    alpha=0.7)
        
        ax2.plot(self.x, self.y, 'ko-', linewidth=2, markersize=8, label='Datos Originales', alpha=0.8)
        
        ax2.set_xlabel('Coordenada x')
        ax2.set_ylabel('Valor y')
        ax2.set_title('Comparación: Datos Originales vs Interpolados')
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        
        plt.tight_layout()
        return fig


class Regresion:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.n = len(x)
    
    def regresion_lineal(self):
        """Calcula regresión lineal y = ax + b"""
        sum_x = sum(self.x)
        sum_y = sum(self.y)
        sum_xy = sum(x_i * y_i for x_i, y_i in zip(self.x, self.y))
        sum_x2 = sum(x_i ** 2 for x_i in self.x)
        
        denominador = self.n * sum_x2 - sum_x ** 2
        a = (self.n * sum_xy - sum_x * sum_y) / denominador
        b = (sum_y * sum_x2 - sum_x * sum_xy) / denominador
        
        y_pred = [a * x_i + b for x_i in self.x]
        stats = self._calcular_estadisticas(y_pred)
        
        return a, b, y_pred, stats
    
    def regresion_polinomial_grado2(self):
        """Calcula regresión polinomial de segundo grado y = ax² + bx + c"""
        sum_x = sum(self.x)
        sum_y = sum(self.y)
        sum_x2 = sum(x_i ** 2 for x_i in self.x)
        sum_x3 = sum(x_i ** 3 for x_i in self.x)
        sum_x4 = sum(x_i ** 4 for x_i in self.x)
        sum_xy = sum(x_i * y_i for x_i, y_i in zip(self.x, self.y))
        sum_x2y = sum((x_i ** 2) * y_i for x_i, y_i in zip(self.x, self.y))
        
        matriz = [
            [self.n, sum_x, sum_x2],
            [sum_x, sum_x2, sum_x3],
            [sum_x2, sum_x3, sum_x4]
        ]
        resultados = [sum_y, sum_xy, sum_x2y]
        
        n_sistema = 3
        for i in range(n_sistema):
            max_row = max(range(i, n_sistema), key=lambda r: abs(matriz[r][i]))
            matriz[i], matriz[max_row] = matriz[max_row], matriz[i]
            resultados[i], resultados[max_row] = resultados[max_row], resultados[i]
            
            for j in range(i + 1, n_sistema):
                factor = matriz[j][i] / matriz[i][i]
                for k in range(i, n_sistema):
                    matriz[j][k] -= factor * matriz[i][k]
                resultados[j] -= factor * resultados[i]
        
        solucion = [0] * n_sistema
        for i in range(n_sistema - 1, -1, -1):
            solucion[i] = resultados[i]
            for j in range(i + 1, n_sistema):
                solucion[i] -= matriz[i][j] * solucion[j]
            solucion[i] /= matriz[i][i]
        
        c, b, a = solucion
        
        y_pred = [a * (x_i ** 2) + b * x_i + c for x_i in self.x]
        stats = self._calcular_estadisticas(y_pred)
        
        return a, b, c, y_pred, stats
    
    def _calcular_estadisticas(self, y_pred):
        """Calcula estadísticos para evaluar el ajuste"""
        media_y = sum(self.y) / self.n
        sst = sum((y_i - media_y) ** 2 for y_i in self.y)
        desv_estandar_total = (sst / (self.n - 1)) ** 0.5
        
        sse = sum((y_real - y_pred_i) ** 2 for y_real, y_pred_i in zip(self.y, y_pred))
        error_estandar = (sse / (self.n - 2)) ** 0.5
        
        sum_xy = sum((x_i - sum(self.x)/self.n) * (y_i - sum(self.y)/self.n) 
                    for x_i, y_i in zip(self.x, self.y))
        sum_x2 = sum((x_i - sum(self.x)/self.n) ** 2 for x_i in self.x)
        sum_y2 = sum((y_i - sum(self.y)/self.n) ** 2 for y_i in self.y)
        
        if sum_x2 == 0 or sum_y2 == 0:
            coef_correlacion = 0
        else:
            coef_correlacion = sum_xy / (sum_x2 * sum_y2) ** 0.5
        
        ssr = sum((y_pred_i - media_y) ** 2 for y_pred_i in y_pred)
        coef_determinacion = ssr / sst if sst != 0 else 0
        
        return {
            'desviacion_estandar_total': desv_estandar_total,
            'error_estandar_estimado': error_estandar,
            'coeficiente_correlacion': coef_correlacion,
            'coeficiente_determinacion': coef_determinacion
        }
    
    def mostrar_resultados(self):
        """Muestra resultados de regresiones"""
        print("\n" + "="*70)
        print("ANÁLISIS DE REGRESIÓN")
        print("="*70)
        
        # Regresión lineal
        a_lineal, b_lineal, y_pred_lineal, stats_lineal = self.regresion_lineal()
        print(f"\n{'─'*70}")
        print("REGRESIÓN LINEAL")
        print(f"{'─'*70}")
        print(f"Ecuación: y = {a_lineal:.4f}x + {b_lineal:.4f}")
        print(f"Desviación estándar total: {stats_lineal['desviacion_estandar_total']:.4f}")
        print(f"Error estándar del estimado: {stats_lineal['error_estandar_estimado']:.4f}")
        print(f"Coeficiente de correlación: {stats_lineal['coeficiente_correlacion']:.4f}")
        print(f"Coeficiente de determinación (R²): {stats_lineal['coeficiente_determinacion']:.4f}")
        
        # Regresión polinomial
        a_poli, b_poli, c_poli, y_pred_poli, stats_poli = self.regresion_polinomial_grado2()
        print(f"\n{'─'*70}")
        print("REGRESIÓN POLINOMIAL GRADO 2")
        print(f"{'─'*70}")
        print(f"Ecuación: y = {a_poli:.4f}x² + {b_poli:.4f}x + {c_poli:.4f}")
        print(f"Desviación estándar total: {stats_poli['desviacion_estandar_total']:.4f}")
        print(f"Error estándar del estimado: {stats_poli['error_estandar_estimado']:.4f}")
        print(f"Coeficiente de correlación: {stats_poli['coeficiente_correlacion']:.4f}")
        print(f"Coeficiente de determinación (R²): {stats_poli['coeficiente_determinacion']:.4f}")
        
        # Comparación
        print(f"\n{'─'*70}")
        print("COMPARACIÓN DE MODELOS")
        print(f"{'─'*70}")
        print(f"Diferencia en R²: {stats_poli['coeficiente_determinacion'] - stats_lineal['coeficiente_determinacion']:.4f}")
        print(f"Diferencia en error estándar: {stats_lineal['error_estandar_estimado'] - stats_poli['error_estandar_estimado']:.4f}")
        
        if stats_poli['coeficiente_determinacion'] > stats_lineal['coeficiente_determinacion']:
            print("➜ El modelo polinomial tiene mejor ajuste (mayor R²)")
        else:
            print("➜ El modelo lineal tiene mejor ajuste (mayor R²)")
        
        return y_pred_lineal, y_pred_poli, stats_lineal, stats_poli
    
    def graficar_resultados(self, y_lineal, y_polinomial, stats_lineal, stats_polinomial):
        """Genera gráfico con los datos y las regresiones"""
        fig = plt.figure(figsize=(12, 8))
        
        plt.scatter(self.x, self.y, color='black', label='Datos originales', alpha=0.7, s=50)
        
        x_ordenado, y_lineal_ordenado = zip(*sorted(zip(self.x, y_lineal)))
        _, y_polinomial_ordenado = zip(*sorted(zip(self.x, y_polinomial)))
        
        plt.plot(x_ordenado, y_lineal_ordenado, 'r-', linewidth=2, 
                label=f'Regresión Lineal (R² = {stats_lineal["coeficiente_determinacion"]:.4f})')
        
        plt.plot(x_ordenado, y_polinomial_ordenado, 'b-', linewidth=2, 
                label=f'Regresión Polinomial Grado 2 (R² = {stats_polinomial["coeficiente_determinacion"]:.4f})')
        
        plt.xlabel('X')
        plt.ylabel('Y')
        plt.title('Comparación de Regresiones Lineal y Polinomial')
        plt.legend()
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        return fig


def buscar_archivo():
    """Busca el archivo datos.xlsx en diferentes ubicaciones"""
    # Obtener la carpeta de Documentos del usuario actual (universal)
    home = os.path.expanduser("~")  # Obtiene la carpeta del usuario (ej: C:/Users/NombreUsuario)
    documentos = os.path.join(home, "Documents")  # Para Windows en inglés
    documentos_es = os.path.join(home, "Documentos")  # Para Windows en español
    
    rutas = [
        # Carpeta actual (donde está el script)
        "./datos.xlsx",
        "./datos.xls",
        "datos.xlsx",
        "datos.xls",
        # Carpeta de Documentos (Windows inglés)
        os.path.join(documentos, "datos.xlsx"),
        os.path.join(documentos, "datos.xls"),
        # Carpeta de Documentos (Windows español)
        os.path.join(documentos_es, "datos.xlsx"),
        os.path.join(documentos_es, "datos.xls"),
        # Escritorio (por si acaso)
        os.path.join(home, "Desktop", "datos.xlsx"),
        os.path.join(home, "Escritorio", "datos.xlsx")
    ]
    
    for ruta in rutas:
        if os.path.exists(ruta):
            print(f"Archivo encontrado: {ruta}")
            return ruta
    
    print("No se encontró el archivo 'datos.xlsx'")
    print("\nRutas verificadas:")
    for ruta in rutas:
        print(f"  - {ruta}")
    print(f"\nSugerencia: Coloca 'datos.xlsx' en tu carpeta de Documentos:")
    print(f"  {documentos}")
    return None


def main():
    print("="*70)
    print(" ANÁLISIS COMPLETO: INTERPOLACIÓN Y REGRESIÓN")
    print("="*70)
    
    # Buscar y cargar archivo
    archivo = buscar_archivo()
    if not archivo:
        print("\nPor favor, asegúrate de que el archivo 'datos.xlsx' existe")
        return
    
    try:
        # Leer datos
        datos = pd.read_excel(archivo)
        
        if 'x' not in datos.columns or 'y' not in datos.columns:
            print("Error: El archivo debe contener columnas 'x' y 'y'")
            print(f"Columnas encontradas: {list(datos.columns)}")
            return
        
        x = datos['x'].tolist()
        y = datos['y'].tolist()
        
        # Limpiar datos
        x_clean = []
        y_clean = []
        for i in range(len(x)):
            if not (pd.isna(x[i]) or pd.isna(y[i])):
                x_clean.append(float(x[i]))
                y_clean.append(float(y[i]))
        
        print(f"\n✓ Datos cargados: {len(x_clean)} puntos")
        print("\nPrimeros 5 puntos:")
        for i in range(min(5, len(x_clean))):
            print(f"  Punto {i+1}: x={x_clean[i]:.2f}, y={y_clean[i]:.2f}")
        
        if len(x_clean) < 3:
            print("Error: Se necesitan al menos 3 puntos")
            return
        
        # ===== ANÁLISIS DE REGRESIÓN =====
        regresion = Regresion(x_clean, y_clean)
        y_lineal, y_poli, stats_lin, stats_pol = regresion.mostrar_resultados()
        
        # ===== INTERPOLACIÓN DE LAGRANGE =====
        interpolacion = InterpolacionLagrange(x_clean, y_clean)
        interpolacion.mostrar_resultados()
        
        # ===== GRÁFICAS =====
        print("\n" + "="*70)
        print("GENERANDO GRÁFICAS...")
        print("="*70)
        
        # Gráfica de regresión
        fig_regresion = regresion.graficar_resultados(y_lineal, y_poli, stats_lin, stats_pol)
        
        # Gráfica de interpolación
        fig_interpolacion = interpolacion.graficar_errores()
        
        plt.show()
        
        print("\nAnálisis completado exitosamente")
        
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()