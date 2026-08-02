import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(
    page_title="Torneo _TC2000 - Dashboard",
    page_icon="🏎️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS oscuros tipo F1 TV
st.markdown("""
    <style>
    .main { background-color: #0f111a; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #161925; }
    div.stButton > button:first-child { background-color: #e10600; color: white; border-radius: 5px; }
    .metric-box {
        background-color: #161925;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #e10600;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# Nombre de tu archivo local
ARCHIVO_EXCEL = "Torneo.xlsx"
# =========================================================================
# 🏎️ GENERADOR DINÁMICO DE OPCIONES PARA LOS DESPLEGABLES (FECHA + CIRCUITO)
# =========================================================================
opciones_fechas_combinadas = ["Campeonato Completo"]

try:
    # Leemos la pestaña sin asumir nombres de columnas (header=None)
    df_crudo_aux = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', header=None, engine='openpyxl')
    
    # Rellenamos hacia abajo la columna A para quitar los huecos de celdas vacías/combinadas
    columna_a_rellenada = df_crudo_aux.iloc[:, 0].ffill()
    
    # Extraemos los nombres únicos en estricto orden de aparición en el Excel
    circuitos_detectados = []
    for c in columna_a_rellenada.dropna().astype(str).str.strip():
        # Filtramos valores basura o nombres de pilotos que puedan colarse en la primera fila
        c_upper = c.upper()
        if c_upper not in [x.upper() for x in circuitos_detectados] and c != "" and "NAN" not in c_upper and not c_upper.startswith("FECHA"):
            circuitos_detectados.append(c)
    
    # Armamos la lista impecable: "Fecha 1 - CABALEN", "Fecha 2 - LA PLATA", etc.
    for idx, nombre_circuito in enumerate(circuitos_detectados):
        opciones_fechas_combinadas.append(f"Fecha {idx + 1} - {nombre_circuito}")
        
except Exception as e:
    # Respaldo de seguridad por si el archivo está bloqueado o da un error inesperado
    opciones_fechas_combinadas = ["Campeonato Completo"] + [f"Fecha {i}" for i in range(1, 11)]
# =========================================================================


# Función para convertir el texto "01:29,228" o "01:31,859" a segundos decimales puros
def tiempo_a_segundos(tiempo_str):
    import datetime
    try:
        if pd.isna(tiempo_str):
            return None
        
        # 1. Si Excel lo guardó como objeto de tiempo nativo (datetime.time)
        if isinstance(tiempo_str, (datetime.time, datetime.datetime)):
            total_segundos = (tiempo_str.minute * 60) + tiempo_str.second + (tiempo_str.microsecond / 1000000.0)
            if total_segundos <= 5.0:
                return None
            return total_segundos
            
        # 2. Si viene como texto tradicional (ej: "01:48,374")
        t_str = str(tiempo_str).replace(',', '.').strip()
        
        if t_str.count(':') == 2:
            partes = t_str.split(':')
            t_str = f"{partes[1]}:{partes[2]}"
            
        if ':' in t_str:
            partes = t_str.split(':')
            minutos = float(partes[0])
            segundos = float(partes[1])
            if minutos == 0 and segundos <= 5.0:
                return None
            return (minutos * 60.0) + segundos
            
        num = float(t_str)
        if num <= 5.0:
            return None
        return num
    except:
        return None

    # Función para dar formato F1 "+0.123" a las diferencias de tiempo
def formato_diferencia(segundos):
    if segundos == 0:
        return "0.000"
    return f"+{segundos:.3f}"

# Función inteligente para leer el archivo local mapeando las hojas reales
def cargar_datos_locales():
    if os.path.exists(ARCHIVO_EXCEL):
        try:
            df_datos = pd.read_excel(ARCHIVO_EXCEL, sheet_name='DATOS', engine='openpyxl')
            df_tiempos = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', engine='openpyxl')


            return df_datos, df_tiempos
        except Exception as e:
            st.error(f"Error al leer el archivo Excel: {e}")
            return None, None
    else:
        st.warning(f"No se encontró el archivo '{ARCHIVO_EXCEL}' en la carpeta del proyecto.")
        return None, None

df_datos, df_tiempos = cargar_datos_locales()

# 2. MENÚ LATERAL
st.sidebar.image("https://flaticon.com", width=80) 
st.sidebar.title("Torneo _TC2000")
st.sidebar.subheader("Campeonato Interno")

opcion = st.sidebar.radio(
    "Navegación",
    ["Resumen", "Posiciones", "Lastre", "Comparativa de Tiempos", "Estadísticas", "Duelo H2H"]
)

# Pilotos oficiales en tu orden exacto de columnas del Excel
pilotos = ["Agus", "Pablo", "Juandi", "Eze"]
df_resumen = pd.DataFrame({
    "Piloto": pilotos, 
    "Puntos":[0, 0, 0, 0], 
    "Victorias":[0, 0, 0, 0], 
    "Poles":[0, 0, 0, 0], 
    "Lastre Actual (kg)": [0, 0, 0, 0]
})

if opcion == "Resumen":
    st.title("📊 Resumen del Campeonato")
    st.write("Sincronizado con tu archivo local Torneo.xlsx")
    
        # Selector de vista unificado y limpio
    if 'opciones_fechas_combinadas' not in locals() and 'opciones_fechas_combinadas' not in globals():
            opciones_fechas_combinadas = ["Campeonato Completo", "Fecha 1", "Fecha 2", "Fecha 3", "Fecha 4", "Fecha 5"]

    vista_seleccionada_combinada = st.selectbox("Seleccionar Fecha o Histórico:", opciones_fechas_combinadas)
        
    if vista_seleccionada_combinada == "Campeonato Completo":
            vista_seleccionada = "Campeonato Completo"
    else:
            vista_seleccionada = str(vista_seleccionada_combinada).split(" - ")[0]


    
    df_resumen = pd.DataFrame({"Piloto": pilotos, "Puntos": [0.0]*4})
    
    if os.path.exists(ARCHIVO_EXCEL):
        try:
            # Cargamos la hoja limpia sin procesar encabezados
            df_hoja1 = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')
            
            # Buscamos los índices de las filas claves en la columna F (índice 5)
            columna_f = df_hoja1.iloc[:, 5].astype(str).str.strip().str.upper()
            indices_totales = columna_f[columna_f == 'TOTAL'].index.tolist()
            indices_total_fecha = columna_f[columna_f == 'TOTAL FECHA'].index.tolist()
            
            # Coordenadas exactas de columnas de pilotos: G=Agus(6), I=Pablo(8), K=Juandi(10), M=Eze(12)
            indices_pilotos = {"Agus": 6, "Pablo": 8, "Juandi": 10, "Eze": 12}
            
            puntos_finales = {p: 0.0 for p in pilotos}
            
            # VISTA A: CAMPEONATO COMPLETO (Suma los TOTAL FECHA de todas las carreras disputadas)
            if vista_seleccionada == "Campeonato Completo":
                for idx_fila in indices_total_fecha:
                    for p, col_idx in indices_pilotos.items():
                        val = pd.to_numeric(df_hoja1.iloc[idx_fila, col_idx], errors='coerce')
                        val_limpio = val if (pd.notna(val) and val > 0) else 0.0
                        puntos_finales[p] += val_limpio  # Vamos acumulando la suma de cada fecha
                        
            # VISTA B: FECHA ESPECÍFICA (Busca el TOTAL FECHA correspondiente al número seleccionado)
            else:
                # Extraemos el número de fecha elegido (ej: "Fecha 7" -> índice 6 en la lista de bloques)
                numero_fecha = int(vista_seleccionada.split()[-1])
                idx_bloque = numero_fecha - 1 # Ajuste a índice 0 de Python
                
                # Validamos que el bloque solicitado exista en el Excel
                if idx_bloque < len(indices_total_fecha):
                    fila_total_fecha = indices_total_fecha[idx_bloque]
                    
                    for p, col_idx in indices_pilotos.items():
                        val = pd.to_numeric(df_hoja1.iloc[fila_total_fecha, col_idx], errors='coerce')
                        puntos_finales[p] = val if pd.notna(val) else 0.0
            
            # Asignamos los puntajes extraídos al DataFrame de visualización
            for p in pilotos:
                df_resumen.loc[df_resumen["Piloto"] == p, "Puntos"] = float(puntos_finales[p])
                    
        except Exception as e:
            st.error(f"Error al procesar los datos de la hoja: {e}")
            
    df_resumen = df_resumen.sort_values(by="Puntos", ascending=False)
    
    # 2. RENDERIZADO DE TABLA Y GRÁFICO
    izq, der = st.columns(2)
    with izq:
        st.subheader(f"🏆 Tabla de Posiciones ({vista_seleccionada})")
        st.dataframe(df_resumen, use_container_width=True, hide_index=True)
    with der:
        st.subheader("📈 Gráfico de Rendimiento")
        fig = px.bar(df_resumen, x="Piloto", y="Puntos", color="Piloto", 
                     title=f"Puntos - {vista_seleccionada}",
                     color_discrete_sequence=["#e10600", "#1f77b4", "#ff7f0e", "#2ca02c"])
        fig.update_layout(template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig, use_container_width=True)
elif opcion == "Comparativa de Tiempos":
    st.title("⏱️ Diferencia de Ritmo y Poles (Histórico vs Fecha)")
    st.write("Análisis estadístico milimétrico basado en tu pestaña 'Carga Tiempos'")

    if df_tiempos is not None:
        # Selector directo tradicional sin solapas intermedias
                # Selector directo tradicional sin solapas intermedias
        fecha_tiempos_sel_combinada = st.selectbox("Seleccionar Período a Analizar:", opciones_fechas_combinadas)
        
        if fecha_tiempos_sel_combinada == "Campeonato Completo":
            fecha_tiempos_sel = "Campeonato Completo"
        else:
            fecha_tiempos_sel = str(fecha_tiempos_sel_combinada).split(" - ")[0]
        
        df_tiempos_crudo = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', header=None, engine='openpyxl')

        
        df_tiempos_crudo = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', header=None, engine='openpyxl')
        df_tiempos_crudo.iloc[:, 0] = df_tiempos_crudo.iloc[:, 0].ffill()
        indices_pilotos_fechas = {"Agus": 2, "Pablo": 3, "Juandi": 4, "Eze": 5}
        columna_circuitos = df_tiempos_crudo.iloc[:, 0].astype(str).str.strip().str.upper()
        columna_sesiones = df_tiempos_crudo.iloc[:, 1].astype(str).str.strip().str.upper()
        
        # DETECCIÓN AUTOMÁTICA DE CIRCUITOS (Sin escribir nombres fijos)
        circuitos_ordenados = []
        for c in df_tiempos_crudo.iloc[:, 0].dropna().astype(str).str.strip():
            if c.upper() not in [x.upper() for x in circuitos_ordenados] and c != "":
                circuitos_ordenados.append(c)

        def buscar_coordenada_fila(nombre_circuito, palabra_sesion_key):
            indices_circuito = columna_circuitos[columna_circuitos == nombre_circuito.upper()].index.tolist()
            if indices_circuito:
                for f in indices_circuito:
                    if f < len(columna_sesiones):
                        if palabra_sesion_key.upper() in str(columna_sesiones.iloc[f]).strip().upper():
                            return f
            return None

        if fecha_tiempos_sel == "Campeonato Completo":
            st.header("📈 Resumen Global Acumulado (Todas las Carreras Disputadas)")
            
            def calcular_promedios_historicos_python(palabra_sesion, titulo_seccion):
                tiempos_acumulados = {p: [] for p in pilotos}
                
                # Iteramos sobre los circuitos reales encontrados automáticamente
                for circuito_id in circuitos_ordenados:
                    fila_exacta = buscar_coordenada_fila(circuito_id, palabra_sesion)
                    if fila_exacta is not None:
                        tiempos_validos_fecha = []
                        for p, col_idx in indices_pilotos_fechas.items():
                            seg_limpio = tiempo_a_segundos(df_tiempos_crudo.iloc[fila_exacta, col_idx])
                            if seg_limpio is not None and seg_limpio > 30.0:
                                tiempos_validos_fecha.append(seg_limpio)
                        if not tiempos_validos_fecha:
                            continue
                        peor_tiempo_fecha = max(tiempos_validos_fecha)
                        tiempo_penalizado = peor_tiempo_fecha * 1.60
                        for p, col_idx in indices_pilotos_fechas.items():
                            seg_limpio = tiempo_a_segundos(df_tiempos_crudo.iloc[fila_exacta, col_idx])
                            if seg_limpio is not None and seg_limpio > 30.0:
                                tiempos_acumulados[p].append(seg_limpio)
                            else:
                                tiempos_acumulados[p].append(tiempo_penalizado)
                
                promedios_finales = {}
                for p in pilotos:
                    if tiempos_acumulados[p]:
                        promedios_finales[p] = sum(tiempos_acumulados[p]) / len(tiempos_acumulados[p])
                
                if promedios_finales:
                    lider_t = min(promedios_finales, key=promedios_finales.get)
                    tiempo_base = promedios_finales[lider_t]
                    tabla_global = []
                    for p in pilotos:
                        # 🛠️ CORRECCIÓN: Validamos de forma segura la existencia del promedio y el tipo de dato
                        if p in promedios_finales and promedios_finales[p] is not None:
                            brecha = promedios_finales[p] - tiempo_base
                            # Si segundos_a_formato o formato_diferencia fallan por alcance, usamos un fallback seguro
                            try:
                                t_formateado = segundos_a_formato(promedios_finales[p])
                            except:
                                t_formateado = f"{promedios_finales[p]:.3f}s"
                                
                            try:
                                b_formateada = "0" if brecha == 0 else formato_diferencia(brecha)
                            except:
                                b_formateada = "0" if brecha == 0 else f"+{brecha:.3f}s"
                                
                            tabla_global.append({"Piloto": p, "Tiempo": t_formateado, "Brecha con Líder": b_formateada, "Orden_Num": brecha})
                        else:
                            tabla_global.append({"Piloto": p, "Tiempo": "-", "Brecha con Líder": "-", "Orden_Num": 999.0})
                    st.subheader(titulo_seccion)
                    st.dataframe(pd.DataFrame(tabla_global).sort_values("Orden_Num")[["Piloto", "Tiempo", "Brecha con Líder"]], use_container_width=True, hide_index=True)

            col1, col2, col3 = st.columns(3)
            with col1: calcular_promedios_historicos_python("CLASIF", "🏎️ Clasificación (Poles)")
            with col2: calcular_promedios_historicos_python("CARRERA 1", "🏁 Carrera 1")
            with col3: calcular_promedios_historicos_python("CARRERA 2", "🏁 Carrera 2")
        else:
            try:
                idx_fecha = int(fecha_tiempos_sel.replace("Fecha ", "")) - 1
                circuito_buscado = circuitos_ordenados[idx_fecha]
            except (IndexError, ValueError):
                circuito_buscado = None

            if circuito_buscado:
                fila_clasif = buscar_coordenada_fila(circuito_buscado, "CLASIF")
                fila_c1 = buscar_coordenada_fila(circuito_buscado, "CARRERA 1")
                fila_c2 = buscar_coordenada_fila(circuito_buscado, "CARRERA 2")
                
                def analizar_sesion_individual_dinamica(fila_real, tipo_sesion):
                    if fila_real is None:
                        st.warning(f"Sin registros para {tipo_sesion} en esta fecha.")
                        return
                    tiempos_pilotos_fecha = {}
                    for p, col_idx in indices_pilotos_fechas.items():
                        seg = tiempo_a_segundos(df_tiempos_crudo.iloc[fila_real, col_idx])
                        if seg is not None and seg > 30.0: tiempos_pilotos_fecha[p] = seg
                    if tiempos_pilotos_fecha:
                        tiempo_base = tiempos_pilotos_fecha[min(tiempos_pilotos_fecha, key=tiempos_pilotos_fecha.get)]
                        tabla_brechas = []
                        for p in pilotos:
                            if p in tiempos_pilotos_fecha:
                                brecha = tiempos_pilotos_fecha[p] - tiempo_base
                                try:
                                    t_formateado = segundos_a_formato(tiempos_pilotos_fecha[p])
                                except:
                                    t_formateado = f"{tiempos_pilotos_fecha[p]:.3f}s"
                                    
                                try:
                                    b_formateada = "0" if brecha == 0 else formato_diferencia(brecha)
                                except:
                                    b_formateada = "0" if brecha == 0 else f"+{brecha:.3f}s"
                                    
                                tabla_brechas.append({"Piloto": p, "Tiempo": t_formateado, "Brecha con Líder": b_formateada, "Orden_Num": brecha})
                            else:
                                tabla_brechas.append({"Piloto": p, "Tiempo": "-", "Brecha con Líder": "-", "Orden_Num": 999.0})
                        st.subheader(f"📋 {tipo_sesion}")
                        st.dataframe(pd.DataFrame(tabla_brechas).sort_values(by="Orden_Num")[["Piloto", "Tiempo", "Brecha con Líder"]], use_container_width=True, hide_index=True)

                col1, col2, col3 = st.columns(3)
                with col1: analizar_sesion_individual_dinamica(fila_clasif, "Clasificación")
                with col2: analizar_sesion_individual_dinamica(fila_c1, "Carrera 1")
                with col3: analizar_sesion_individual_dinamica(fila_c2, "Carrera 2")
            else:
                st.warning(f"La {fecha_tiempos_sel} aún no tiene datos cargados en el Excel.")

elif opcion == "Lastre":
    st.title("⚖️ Control de Lastre Oficial")
    st.write("Historial y penalizaciones por kilogramos en pista según reglamento TC2000.")

    if os.path.exists(ARCHIVO_EXCEL):
        try:
            # Leemos la Hoja1 de forma nativa sin procesar encabezados
            df_hoja1 = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')
            
            # Convertimos la columna F (índice 5) a texto limpio para buscar las filas clave
            columna_f = df_hoja1.iloc[:, 5].astype(str).str.strip().str.upper()
            
            # Encontramos TODAS las filas del Excel que contienen "LASTRE FECHA" y "LASTRE ACUMULADO"
            filas_lastre_fecha = columna_f[columna_f == "LASTRE FECHA"].index.tolist()
            filas_lastre_acum = columna_f[columna_f == "LASTRE ACUMULADO"].index.tolist()
            
            # Mapeo exacto de columnas de pilotos en la Hoja1: G=Agus(6), I=Pablo(8), K=Juandi(10), M=Eze(12)
            indices_pilotos_lastre = {"Agus": 6, "Pablo": 8, "Juandi": 10, "Eze": 12}
            
            # El selector va a tener tantas Fechas como bloques tengas cargados en el Excel
            total_fechas_detectadas = len(filas_lastre_fecha)
            if total_fechas_detectadas > 0:
                opciones_lastre_dinamicas = [opt for opt in opciones_fechas_combinadas if opt != "Campeonato Completo"][:total_fechas_detectadas]
                fecha_sel = st.selectbox("Seleccionar Fecha para Consultar Lastre Técnico:", opciones_lastre_dinamicas)
                
                # 🛠️ EXTRACCIÓN INFALIBLE: Sacamos solo el número sin que le afecte el nombre del circuito
                import re
                numeros_encontrados = re.findall(r'\d+', str(fecha_sel))
        # 🛠️ EXTRAE EL NÚMERO DE FORMA INFALIBLE
                import re
                numeros_encontrados = re.findall(r'\d+', str(fecha_sel))
                idx_fecha_sel = int(numeros_encontrados[0]) - 1 if numeros_encontrados else 0
        
        # [AQUÍ ABAJO DEBE CONTINUAR TU LÓGICA ORIGINAL DE LA TABLA: tabla_lastre = [] ]

                tabla_lastre = []
                
                for piloto, col_idx in indices_pilotos_lastre.items():
                    # 1. LASTRE GENERADO EN LA FECHA ACTUAL
                    # Vamos directo a la fila "LASTRE FECHA" de la fecha seleccionada
                    fila_generado = filas_lastre_fecha[idx_fecha_sel]
                    val_actual = str(df_hoja1.iloc[fila_generado, col_idx]).upper().replace("KG", "").strip()
                    lastre_generado_en_fecha = pd.to_numeric(val_actual, errors='coerce')
                    if pd.isna(lastre_generado_en_fecha): lastre_generado_en_fecha = 0.0
                    
                    # 2. LASTRE CON EL QUE SE CORRIÓ LA FECHA
                    # Es el "LASTRE ACUMULADO" del bloque de la fecha ANTERIOR.
                    if idx_fecha_sel > 0:
                        fila_acum_anterior = filas_lastre_acum[idx_fecha_sel - 1]
                        val_anterior = str(df_hoja1.iloc[fila_acum_anterior, col_idx]).upper().replace("KG", "").strip()
                        lastre_con_el_que_se_corrio = pd.to_numeric(val_anterior, errors='coerce')
                        if pd.isna(lastre_con_el_que_se_corrio): lastre_con_el_que_se_corrio = 0.0
                    else:
                        lastre_con_el_que_se_corrio = 0.0  # En la Fecha 1 todos largan limpios con 0 kg
                    
                    tabla_lastre.append({
                        "Piloto": piloto,
                        "Lastre en Pista (kg)": f"{int(lastre_con_el_que_se_corrio)} kg",
                        "Lastre Generado (kg)": f"+{int(lastre_generado_en_fecha)} kg" if lastre_generado_en_fecha > 0 else "0 kg",
                        "Orden_Visual": lastre_con_el_que_se_corrio
                    })
                
                # Ordenamos de mayor a menor peso en pista para las tarjetas
                df_lastre_render = pd.DataFrame(tabla_lastre).sort_values(by="Orden_Visual", ascending=False)
                
                st.subheader(f"📊 Estado de Penalizaciones — {fecha_sel}")
                
                # Renderizado de Tarjetas F1
                cols_cards = st.columns(4)
                for i, (_, row) in enumerate(df_lastre_render.iterrows()):
                    with cols_cards[i % 4]:
                        st.markdown(f"""
                        <div class="metric-box">
                           <h4 style='margin:0; color:#e10600;'>{row['Piloto']}</h4>
                           <p style='margin:5px 0 0 0; font-size:24px; font-weight:bold;'>{row['Lastre en Pista (kg)']}</p>
                           <small style='color:#aaa;'>Generó en carrera: {row['Lastre Generado (kg)']}</small>
                        </div>
                        """, unsafe_allow_html=True)
                
                st.markdown("---")
                st.dataframe(df_lastre_render[["Piloto", "Lastre en Pista (kg)", "Lastre Generado (kg)"]], use_container_width=True, hide_index=True)
            else:
                st.warning("No se detectaron las filas de 'Lastre Fecha' en la 'Hoja1'. Revisá que estén escritas exactamente igual.")
                
        except Exception as e:
            st.error(f"Error al calcular el lastre desde el Excel: {e}")
    else:
        st.warning(f"No se encontró el archivo '{ARCHIVO_EXCEL}'.")
elif opcion == "Posiciones":
    st.title("🏆 Centro de Cómputos del Campeonato")
    st.write("Análisis interactivo de puntajes por fecha y evolución de la tabla general.")

    if os.path.exists(ARCHIVO_EXCEL):
        try:
            # Leemos la Hoja1 de forma nativa sin procesar encabezados
            df_hoja1 = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')
            
            # Convertimos la columna F a texto limpio para mapear las filas clave
            columna_f = df_hoja1.iloc[:, 5].astype(str).str.strip().str.upper()
            
            # Buscamos de forma estricta todas las filas que dicen exactamente "TOTAL FECHA"
            filas_total_fecha = columna_f[columna_f == "TOTAL FECHA"].index.tolist()
            
            # Coordenadas de columnas de pilotos: G=Agus(6), I=Pablo(8), K=Juandi(10), M=Eze(12)
            indices_pilotos_pos = {"Agus": 6, "Pablo": 8, "Juandi": 10, "Eze": 12}
            
            total_fechas_detectadas = len(filas_total_fecha)
            
            if total_fechas_detectadas > 0:
                # 1. SELECTOR INTERACTIVO DE FECHAS
                opciones_fechas = [f"Fecha {i}" for i in range(1, total_fechas_detectadas + 1)]
                fecha_seleccionada = st.selectbox("Seleccionar Fecha a Consultar:", opciones_fechas, index=total_fechas_detectadas-1)
                idx_fecha_sel = int(fecha_seleccionada.split()[-1]) - 1  
                
                # Fila base de "TOTAL FECHA" para la carrera elegida
                fila_total_fecha_actual = filas_total_fecha[idx_fecha_sel]
                # Fila de "TOTAL" acumulado (siempre 2 renglones más abajo)
                fila_total_acumulado = fila_total_fecha_actual + 2
                
                tabla_fecha_pura = []
                tabla_campeonato_historico = []
                
                for piloto, col_idx in indices_pilotos_pos.items():
                    puntos_fecha_puros = 0.0
                    puntos_acumulados_historicos = 0.0
                    
                    # A. Leemos los puntos netos de la fecha (Limpiando textos ocultos de Excel)
                    if fila_total_fecha_actual < len(df_hoja1):
                        val_fecha = str(df_hoja1.iloc[fila_total_fecha_actual, col_idx]).replace(',', '.').strip()
                        puntos_fecha_puros = pd.to_numeric(val_fecha, errors='coerce')
                        if pd.isna(puntos_fecha_puros): puntos_fecha_puros = 0.0
                    
                    # B. Leemos los puntos del TOTAL acumulado (Forzando la conversión matemática)
                    if fila_total_acumulado < len(df_hoja1):
                        val_acum = str(df_hoja1.iloc[fila_total_acumulado, col_idx]).replace(',', '.').strip()
                        puntos_acumulados_historicos = pd.to_numeric(val_acum, errors='coerce')
                        if pd.isna(puntos_acumulados_historicos): puntos_acumulados_historicos = 0.0
                        
                    tabla_fecha_pura.append({"Piloto": piloto, "Puntos de la Fecha": puntos_fecha_puros})
                    tabla_campeonato_historico.append({"Piloto": piloto, "Puntos Totales": puntos_acumulados_historicos})
                
                # Armamos y ordenamos los dos DataFrames
                df_fecha_ordenado = pd.DataFrame(tabla_fecha_pura).sort_values(by="Puntos de la Fecha", ascending=False).reset_index(drop=True)
                df_campeonato_ordenado = pd.DataFrame(tabla_campeonato_historico).sort_values(by="Puntos Totales", ascending=False).reset_index(drop=True)
                
                # 2. DISEÑO VISUAL EN LA APP (Dividido en dos columnas principales)
                izq, der = st.columns(2)
                
                with izq:
                    st.subheader(f"🏁 Clasificador de la {fecha_seleccionada}")
                    st.write("Puntaje neto obtenido únicamente en este circuito.")
                    st.dataframe(df_fecha_ordenado, use_container_width=True, hide_index=True)
                    
                    fig_fecha = px.bar(df_fecha_ordenado, x="Piloto", y="Puntos de la Fecha", color="Piloto",
                                       color_discrete_sequence=["#e10600", "#1f77b4", "#ff7f0e", "#2ca02c"])
                    fig_fecha.update_layout(template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig_fecha, use_container_width=True, key="grafico_fecha_pura")
                    
                with der:
                    st.subheader(f"🏆 Posiciones Generales (A la {fecha_seleccionada})")
                    st.write("Tabla acumulada histórica con el cierre de esta carrera.")
                    st.dataframe(df_campeonato_ordenado, use_container_width=True, hide_index=True)
                    
                    fig_camp = px.bar(df_campeonato_ordenado, x="Piloto", y="Puntos Totales", color="Piloto",
                                      color_discrete_sequence=["#e10600", "#1f77b4", "#ff7f0e", "#2ca02c"])
                    fig_camp.update_layout(template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig_camp, use_container_width=True, key="grafico_campeonato_acumulado")
                    
            else:
                st.warning("No se detectaron las celdas de 'TOTAL FECHA' en tu 'Hoja1'.")
        except Exception as e:
            st.error(f"Error al procesar el centro de cómputos: {e}")
    else:
        st.warning(f"No se encontró el archivo '{ARCHIVO_EXCEL}'.")

elif opcion == "Estadísticas":
    st.title("📊 Salón de la Fama e Historial Estadístico")
    st.write("Análisis gráfico acumulado del desempeño de todos los pilotos en el torneo.")

    if df_tiempos is not None:
        try:
            # 1. LECTURA DE DATOS PARA POLES Y VUELTAS RÁPIDAS
            df_tiempos_crudo = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', header=None, engine='openpyxl')
            df_tiempos_crudo.iloc[:, 0] = df_tiempos_crudo.iloc[:, 0].ffill()
            columna_sesiones = df_tiempos_crudo.iloc[:, 1].astype(str).str.strip().str.upper()
            indices_pilotos_fechas = {"Agus": 2, "Pablo": 3, "Juandi": 4, "Eze": 5}
            
            # Conteo de Poles
            filas_clasif = df_tiempos_crudo[columna_sesiones.str.contains("CLASIF", na=False)]
            conteo_poles = {p: 0 for p in pilotos}
            for _, fila in filas_clasif.iterrows():
                tiempos_fila = {}
                for p, col_idx in indices_pilotos_fechas.items():
                    seg = tiempo_a_segundos(fila[col_idx])
                    if seg is not None and seg > 30.0: tiempos_fila[p] = seg
                if tiempos_fila:
                    conteo_poles[min(tiempos_fila, key=tiempos_fila.get)] += 1
            df_poles = pd.DataFrame(list(conteo_poles.items()), columns=["Piloto", "Poles Metidas"])
            
            # Conteo de Vueltas Rápidas
            filas_carreras = df_tiempos_crudo[columna_sesiones.str.contains("CARRERA", na=False)]
            conteo_victorias = {p: 0 for p in pilotos}
            for _, fila in filas_carreras.iterrows():
                tiempos_fila = {}
                for p, col_idx in indices_pilotos_fechas.items():
                    seg = tiempo_a_segundos(fila[col_idx])
                    if seg is not None and seg > 30.0: tiempos_fila[p] = seg
                if tiempos_fila:
                    conteo_victorias[min(tiempos_fila, key=tiempos_fila.get)] += 1
            df_vics = pd.DataFrame(list(conteo_victorias.items()), columns=["Piloto", "Sesiones Lideradas"])
            # 2. MOTOR DE RITMO HISTÓRICO (PROCESAMIENTO INDEPENDIENTE DE CARRERA 1 Y 2)
            ritmo_acumulado = {p: [] for p in pilotos}
            
            if os.path.exists(ARCHIVO_EXCEL):
                df_ritmo_crudo = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Diferencia en Carrera', header=None, engine='openpyxl')
                
                # Columnas reales basadas en tus fotos: Carrera 1 (B=1, D=3, F=5, H=7) y Carrera 2 (L=11, N=13, P=15, R=17)
                columnas_c1 = {"Pablo": 1, "Eze": 3, "Agus": 5, "Juandi": 7}   
                columnas_c2 = {"Pablo": 11, "Eze": 13, "Agus": 15, "Juandi": 17} 
                
                filas_totales_hoja, columnas_totales_hoja = df_ritmo_crudo.shape
                
                # Recorremos el Excel haciendo saltos de 24 en 24 filas hacia abajo para las 10 fechas
                for fecha_idx in range(10):
                    fila_base_bloque = fecha_idx * 24
                    fila_ritmo_celda = fila_base_bloque + 22 # Fila 23 física de Excel (Índice 22 de Python)
                    
                    if fila_ritmo_celda < filas_totales_hoja:
                        
                        # --- PROCESAMOS ÚNICAMENTE CARRERA 1 ---
                        lider_c1 = None
                        t_lider_c1 = None
                        
                        for p, col in columnas_c1.items():
                            if col < columnas_totales_hoja:
                                celda = df_ritmo_crudo.iloc[fila_ritmo_celda, col]
                                texto = str(celda).strip().replace(" ", "")
                                if ":" in texto and "+" not in texto:
                                    seg = tiempo_a_segundos(celda)
                                    if seg is not None and seg > 30.0:
                                        lider_c1 = p
                                        t_lider_c1 = seg
                                        break
                        
                        if t_lider_c1 is not None:
                            for p, col in columnas_c1.items():
                                if col < columnas_totales_hoja:
                                    celda = df_ritmo_crudo.iloc[fila_ritmo_celda, col]
                                    texto = str(celda).strip().replace(" ", "")
                                    if p == lider_c1:
                                        ritmo_acumulado[p].append(t_lider_c1)
                                    elif "+" in texto:
                                        try:
                                            brecha = float(texto.replace("+", "").replace(",", "."))
                                            ritmo_acumulado[p].append(t_lider_c1 + brecha)
                                        except: pass

                        # --- PROCESAMOS ÚNICAMENTE CARRERA 2 (LA QUE TIENE LOS DATOS ACTUALES) ---
                        lider_c2 = None
                        t_lider_c2 = None
                        
                        for p, col in columnas_c2.items():
                            if col < columnas_totales_hoja:
                                celda = df_ritmo_crudo.iloc[fila_ritmo_celda, col]
                                texto = str(celda).strip().replace(" ", "")
                                if ":" in texto and "+" not in texto:
                                    seg = tiempo_a_segundos(celda)
                                    if seg is not None and seg > 30.0:
                                        lider_c2 = p
                                        t_lider_c2 = seg
                                        break
                        
                        if t_lider_c2 is not None:
                            for p, col in columnas_c2.items():
                                if col < columnas_totales_hoja:
                                    celda = df_ritmo_crudo.iloc[fila_ritmo_celda, col]
                                    texto = str(celda).strip().replace(" ", "")
                                    if p == lider_c2:
                                        ritmo_acumulado[p].append(t_lider_c2)
                                    elif "+" in texto:
                                        try:
                                            brecha = float(texto.replace("+", "").replace(",", "."))
                                            ritmo_acumulado[p].append(t_lider_c2 + brecha)
                                        except: pass

            # Filtramos los promedios de los pilotos que SÍ tienen datos registrados
            promedios_netos = {}
            for p in pilotos:
                if ritmo_acumulado[p]:
                    promedios_netos[p] = sum(ritmo_acumulado[p]) / len(ritmo_acumulado[p])
            
            # CALCULAMOS LAS BRECHAS HISTÓRICAS GENERALES CON RESPECTO AL MÁS RÁPIDO
            tabla_ritmo_brechas = []
            if promedios_netos:
                piloto_mas_rapido_campeonato = min(promedios_netos, key=promedios_netos.get)
                tiempo_base_campeonato = promedios_netos[piloto_mas_rapido_campeonato]
                
                for p in pilotos:
                    if p in promedios_netos:
                        brecha_global = promedios_netos[p] - tiempo_base_campeonato
                        formato_texto = "Líder (0,000s)" if brecha_global == 0 else f"+{brecha_global:.3f}s".replace(".", ",")
                        
                        tabla_ritmo_brechas.append({
                            "Piloto": p,
                            "Brecha Promedio (seg)": brecha_global,
                            "Texto Visual": formato_texto
                        })
            
            df_ritmo_final = pd.DataFrame(tabla_ritmo_brechas)
            if not df_ritmo_final.empty:
                df_ritmo_final = df_ritmo_final.sort_values(by="Brecha Promedio (seg)", ascending=True)


            # 3. RENDERIZADO GRÁFICO EN PANTALLA
            izq, der = st.columns(2)
            with izq:
                st.subheader("🏎️ Récord de Pole Positions (Sábados)")
                fig_poles = px.pie(df_poles, values="Poles Metidas", names="Piloto", hole=0.4,
                                   color="Piloto", color_discrete_sequence=["#e10600", "#1f77b4", "#ff7f0e", "#2ca02c"])
                fig_poles.update_layout(template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_poles, use_container_width=True, key="grafico_poles_historico")
                
            with der:
                st.subheader("🏁 Dueños de la Vuelta Rápida (Domingos)")
                st.write("Cada fecha otorga 2 récords de vuelta (Carrera 1 y 2). Conteo histórico:")
                fig_vics = px.bar(df_vics, x="Piloto", y="Sesiones Lideradas", color="Piloto",
                                  color_discrete_sequence=["#e10600", "#1f77b4", "#ff7f0e", "#2ca02c"])
                fig_vics.update_traces(texttemplate='<b>%{y} VR</b>', textposition='outside', textfont_size=14)
                fig_vics.update_layout(
                    template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    yaxis_title="Cantidad de Vueltas Rápidas", xaxis_title="Pilotos del Torneo",
                                        yaxis=dict(range=[0, max(df_vics["Sesiones Lideradas"]) + 2 if not df_vics.empty else 5])
                )
                st.plotly_chart(fig_vics, use_container_width=True, key="grafico_victorias_ritmo")
            
            # FILA INFERIOR: GRÁFICO DE RITMO HISTÓRICO POR BRECHAS
            st.markdown("---")
            st.subheader("⏱️ Brecha de Ritmo Total en Carrera (Histórico Acumulado)")
            st.write("Muestra cuántos segundos por vuelta pierde cada piloto en promedio con respecto al líder de ritmo del campeonato.")
            
            if not df_ritmo_final.empty:
                fig_ritmo = px.bar(df_ritmo_final, x="Brecha Promedio (seg)", y="Piloto", color="Piloto",
                                    orientation='h', text="Texto Visual",
                                    color_discrete_sequence=["#2ca02c", "#e10600", "#1f77b4", "#ff7f0e"])
                fig_ritmo.update_traces(textposition='inside', textfont_size=14, textfont_color="white")
                fig_ritmo.update_layout(
                    template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    xaxis_title="Brecha Promedio por Vuelta (Segundos)", yaxis_title="Pilotos",
                    showlegend=False
                )
                st.plotly_chart(fig_ritmo, use_container_width=True, key="grafico_ritmo_total_carrera")
            else:
                st.info("Aún no hay datos de ritmo cargados en la pestaña 'Diferencia en Carrera' para procesar.")
        except Exception as e:
            st.error(f"Error al procesar el historial estadístico: {e}")
    else:
        st.warning(f"No se encontró la pestaña de tiempos para procesar.")

elif opcion == "Duelo H2H":
    st.title("⚔️ Duelo Head-to-Head (H2H)")
    st.write("Comparativa directa y cara a cara entre dos pilotos del torneo")

    # 1. LEER DATOS DESDE LAS DOS PESTAÑAS DEL EXCEL
    try:
        df_puntos_h2h = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Tabla Final', engine='openpyxl')
        df_tiempos_crudo = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', header=None, engine='openpyxl')
    except Exception as e:
        st.error(f"❌ Error al leer el archivo Excel: {e}")
        df_puntos_h2h = None

    if df_puntos_h2h is not None and df_tiempos_crudo is not None:
        df_puntos_h2h.columns = [str(c).strip() for c in df_puntos_h2h.columns]
        
        if "PILOTO" in df_puntos_h2h.columns and "PTS" in df_puntos_h2h.columns:
            df_puntos_h2h = df_puntos_h2h.dropna(subset=["PILOTO"])
            lista_pilotos = sorted(df_puntos_h2h["PILOTO"].unique())
            
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                piloto1 = st.selectbox("Selecciona Piloto 1:", lista_pilotos, index=0, key="h2h_p1")
            with col_p2:
                idx_p2 = 1 if len(lista_pilotos) > 1 else 0
                piloto2 = st.selectbox("Selecciona Piloto 2:", lista_pilotos, index=idx_p2, key="h2h_p2")

            if piloto1 == piloto2:
                st.warning("⚠️ Selecciona dos pilotos diferentes para comparar sus estadísticas.")
            else:
                st.markdown("---")
                
                # --- PROCESAMIENTO DE PUNTOS (Tabla Final) ---
                fila_p1 = df_puntos_h2h[df_puntos_h2h["PILOTO"] == piloto1]
                fila_p2 = df_puntos_h2h[df_puntos_h2h["PILOTO"] == piloto2]
                
                # 🛠️ CORRECCIÓN QUIRÚRGICA: Agregamos [0] para sacar el número real de la lista de NumPy
                puntos_p1 = float(fila_p1["PTS"].values[0]) if not fila_p1.empty else 0.0
                puntos_p2 = float(fila_p2["PTS"].values[0]) if not fila_p2.empty else 0.0
                
                # --- METRICAS DE PUNTOS PRINCIPALES ---
                st.subheader("🏆 Puntos en el Campeonato")
                metric_col1, metric_col2, metric_col3 = st.columns(3)
                with metric_col1:
                    st.metric(label=f"Puntos {piloto1}", value=f"{puntos_p1:.2f} pts")
                with metric_col2:
                    diferencia = abs(puntos_p1 - puntos_p2)
                    quien_gana = piloto1 if puntos_p1 > puntos_p2 else piloto2
                    if puntos_p1 == puntos_p2:
                        st.metric(label="Brecha", value="Empate", delta="0 pts")
                    else:
                        st.metric(label="Brecha", value=f"{diferencia:.2f} pts", delta=f"Lidera {quien_gana}")
                with metric_col3:
                    st.metric(label=f"Puntos {piloto2}", value=f"{puntos_p2:.2f} pts")

                # --- PROCESAMIENTO AUTOMÁTICO DE TIEMPOS (Poles y VRs) ---
                df_tiempos_crudo.iloc[:, 0] = df_tiempos_crudo.iloc[:, 0].ffill()
                indices_pilotos_fechas = {"Agus": 2, "Pablo": 3, "Juandi": 4, "Eze": 5}
                columna_circuitos = df_tiempos_crudo.iloc[:, 0].astype(str).str.strip().str.upper()
                columna_sesiones = df_tiempos_crudo.iloc[:, 1].astype(str).str.strip().str.upper()
                
                circuitos_ordenados = []
                for c in df_tiempos_crudo.iloc[:, 0].dropna().astype(str).str.strip():
                    if c.upper() not in [x.upper() for x in circuitos_ordenados] and c != "":
                        circuitos_ordenados.append(c)

                def buscar_coordenada_fila(nombre_circuito, palabra_sesion_key):
                    indices_circuito = columna_circuitos[columna_circuitos == nombre_circuito.upper()].index.tolist()
                    if indices_circuito:
                        for f in indices_circuito:
                            if f < len(columna_sesiones):
                                if palabra_sesion_key.upper() in str(columna_sesiones.iloc[f]).strip().upper():
                                    return f
                    return None

                poles_p1, poles_p2 = 0, 0
                vr1_p1, vr1_p2 = 0, 0
                vr2_p1, vr2_p2 = 0, 0

                for circuito in circuitos_ordenados:
                    for sesion, tipo in [("CLASIF", "pole"), ("CARRERA 1", "vr1"), ("CARRERA 2", "vr2")]:
                        f_idx = buscar_coordenada_fila(circuito, sesion)
                        if f_idx is not None:
                            try:
                                col_idx_p1 = indices_pilotos_fechas[piloto1]
                                col_idx_p2 = indices_pilotos_fechas[piloto2]
                                
                                t_p1 = tiempo_a_segundos(df_tiempos_crudo.iloc[f_idx, col_idx_p1])
                                t_p2 = tiempo_a_segundos(df_tiempos_crudo.iloc[f_idx, col_idx_p2])
                                
                                if t_p1 is not None and t_p2 is not None and t_p1 > 30.0 and t_p2 > 30.0:
                                    if t_p1 < t_p2:
                                        if tipo == "pole": poles_p1 += 1
                                        elif tipo == "vr1": vr1_p1 += 1
                                        elif tipo == "vr2": vr2_p1 += 1
                                    elif t_p2 < t_p1:
                                        if tipo == "pole": poles_p2 += 1
                                        elif tipo == "vr1": vr1_p2 += 1
                                        elif tipo == "vr2": vr2_p2 += 1
                            except KeyError:
                                pass

                # --- INTERFAZ VISUAL: TARJETAS PERSONALIZADAS MINIMALISTAS ---
                st.markdown("---")
                st.subheader("🏁 Historial de Duelos Directos en Pista")
                st.write("Frecuencia acumulada de quién superó a quién en cada sesión:")

                estilo_tarjeta = """
                <style>
                .tarjeta-f1 {
                    background-color: #1a1c23;
                    border-left: 5px solid #ff1801;
                    padding: 20px;
                    border-radius: 8px;
                    margin-bottom: 15px;
                    box-shadow: 2px 2px 10px rgba(0,0,0,0.3);
                }
                .titulo-tarjeta {
                    color: #ff1801;
                    font-size: 18px;
                    font-weight: bold;
                    margin-bottom: 12px;
                }
                .marcador-tarjeta {
                    color: #ffffff;
                    font-size: 32px;
                    font-weight: bold;
                    font-family: 'monospace';
                    margin-bottom: 0px;
                }
                </style>
                """
                st.markdown(estilo_tarjeta, unsafe_allow_html=True)

                c_pole, c_vr1, c_vr2 = st.columns(3)
                
                with c_pole:
                    st.markdown(f"""
                        <div class="tarjeta-f1">
                            <div class="titulo-tarjeta">⏱️ Clasificación</div>
                            <div class="marcador-tarjeta">{poles_p1} — {poles_p2}</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                with c_vr1:
                    st.markdown(f"""
                        <div class="tarjeta-f1">
                            <div class="titulo-tarjeta">🏎️ VR Carrera 1</div>
                            <div class="marcador-tarjeta">{vr1_p1} — {vr1_p2}</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                with c_vr2:
                    st.markdown(f"""
                        <div class="tarjeta-f1">
                            <div class="titulo-tarjeta">🏁 VR Carrera 2</div>
                            <div class="marcador-tarjeta">{vr2_p1} — {vr2_p2}</div>
                        </div>
                    """, unsafe_allow_html=True)
                    
        else:
            st.error("❌ No se encontraron las columnas exactas 'PILOTO' o 'PTS' en la pestaña 'Tabla Final'.")
