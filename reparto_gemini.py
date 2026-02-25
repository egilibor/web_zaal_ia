import pandas as pd
import os
import re

def limpiar_texto(texto):
    if not isinstance(texto, str): return ""
    texto = texto.upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U').replace('Ñ','N')
    return re.sub(r'[^A-Z0-9 ]', '', texto).strip()

def obtener_columna(df, palabras_clave):
    for col in df.columns:
        col_norm = limpiar_texto(str(col))
        for p in palabras_clave:
            if p in col_norm: return col
    return None

def procesar():
    print("--- 🚚 Motor Logístico: Salida desde Vall d'Uixó (Versión con Rangos) ---")
    archivo = "salida.xlsx"
    if not os.path.exists(archivo):
        print(f"❌ Error: No se encuentra '{archivo}'")
        return

    xl = pd.ExcelFile(archivo)
    # Filtramos hojas de resumen o zonas excluidas
    hojas_disp = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["VINAROZ", "MORELLA", "RESUMEN"])]
    
    print("\nRutas disponibles en el Excel:")
    for i, h in enumerate(hojas_disp):
        print(f"[{i}] {h}")
    
    seleccion = input("\nIntroduce los números o rangos (ej: 0, 1, 3-5): ")
    try:
        indices = []
        # Dividimos por comas primero
        partes = [x.strip() for x in seleccion.split(',')]
        for parte in partes:
            if '-' in parte:
                # Si hay un guion, calculamos el rango
                inicio, fin = parte.split('-')
                indices.extend(range(int(inicio), int(fin) + 1))
            else:
                # Si es un número solo
                indices.append(int(parte))
        
        # Eliminamos duplicados y ordenamos
        indices = sorted(list(set(indices)))
        # Filtramos para no salirnos del índice de hojas disponibles
        hojas_a_procesar = [hojas_disp[i] for i in indices if i < len(hojas_disp)]
        
        if not hojas_a_procesar:
            print("❌ No se seleccionó ninguna ruta válida.")
            return
            
    except Exception as e:
        print(f"❌ Selección no válida. Error: {e}")
        return

    plan_final = []

    for nombre_hoja in hojas_a_procesar:
        print(f"⌛ Procesando {nombre_hoja}...")
        df = pd.read_excel(xl, sheet_name=nombre_hoja)
        if df.empty: continue

        c_dir = obtener_columna(df, ['DIRECCION', 'DIR'])
        c_kgs = obtener_columna(df, ['KILO', 'KGS', 'PESO'])
        c_pob = obtener_columna(df, ['POBLACION', 'POB'])
        
        if not c_dir or not c_kgs:
            print(f"⚠️ Columnas críticas no encontradas en {nombre_hoja}")
            continue

        # Capacidad según tipo de ruta
        capacidad_max = 3500 if any(x in nombre_hoja.upper() for x in ["HOSPITAL", "FEDERAC"]) else 800
        
        # Ordenación geográfica (Sur a Norte)
        df = df.sort_values(by=[c_pob, c_dir])

        viaje_n = 1
        while df[df[c_kgs] > 0].shape[0] > 0:
            peso_act = 0
            paradas = []
            algo_cargado = False
            
            for idx, row in df.iterrows():
                peso_linea = row[c_kgs]
                if peso_linea <= 0: continue
                
                if peso_linea > capacidad_max and peso_act == 0:
                    paradas.append(row)
                    df.at[idx, c_kgs] = 0
                    algo_cargado = True
                    break 
                
                if peso_act + peso_linea <= capacidad_max:
                    paradas.append(row)
                    peso_act += peso_linea
                    df.at[idx, c_kgs] = 0
                    algo_cargado = True
            
            if not algo_cargado: break 

            if paradas:
                res = pd.DataFrame(paradas)
                tag_v = f"{nombre_hoja[:15].strip()}_V{viaje_n}"
                res['VEHICULO'] = tag_v
                plan_final.append(res)
                viaje_n += 1

    if plan_final:
        # NOMBRADO DINÁMICO DEL ARCHIVO
        etiqueta_archivo = "_".join([h[:12].strip().replace(" ", "_") for h in hojas_a_procesar])
        # Limitamos el nombre del archivo si es muy largo
        if len(etiqueta_archivo) > 100:
            etiqueta_archivo = f"VARIAS_RUTAS_{len(hojas_a_procesar)}"
            
        nombre_salida = f"PLAN_{etiqueta_archivo}.xlsx"
        
        with pd.ExcelWriter(nombre_salida, engine='openpyxl') as writer:
            resultado = pd.concat(plan_final)
            for v in resultado['VEHICULO'].unique():
                df_v = resultado[resultado['VEHICULO'] == v].copy()
                df_v.insert(0, 'PARADA', range(1, len(df_v) + 1))
                df_v.to_excel(writer, sheet_name=str(v)[:31], index=False)
        
        print(f"\n✅ Proceso completado con éxito.")
        print(f"📂 Archivo generado: {nombre_salida}")
    else:
        print("\n❌ No se encontraron datos válidos.")

if __name__ == "__main__":
    procesar()