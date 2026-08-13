import streamlit as st
import pandas as pd
import plotly.express as px
import os
import datetime

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
# FUNCIONES AUXILIARES DE TIEMPOS
# =========================================================================

def tiempo_a_segundos(tiempo_str):
    try:
        if pd.isna(tiempo_str):
            return None
        
        # 1. Si Excel lo guardó como objeto de tiempo nativo
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

def formato_diferencia(segundos):
    if segundos == 0:
        return "0.000"
    return f"+{segundos:.3f}"

# =========================================================================
# PROCESADOR DINÁMICO DE DATOS
# =========================================================================

def procesar_hoja_dinamica(file_path, sheet_name):
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
    registros = []
    circuito_actual = "Desconocido"

    for row_idx in range(len(df_raw)):
        row_vals = df_raw.iloc[row_idx].dropna().tolist()
        
        # Detectar títulos de Circuitos
        if len(row_vals) == 1 and isinstance(row_vals[0], str):
            val = row_vals[0].strip()
            if not val.isdigit() and "PILOTO" not in val.upper() and ":" not in val:
                circuito_actual = val
                continue
        
        # Detectar columnas de pilotos y tiempos
        for col_idx in range(len(df_raw.columns) - 1):
            val_header = str(df_raw.iloc[row_idx, col_idx]).strip()
            
            if val_header and val_header.lower() != 'nan' and not val_header.isdigit() and val_header.upper() != 'PILOTO':
                for r in range(row_idx + 1, len(df_raw)):
                    vuelta_num = df_raw.iloc[r, col_idx]
                    tiempo_val = df_raw.iloc[r, col_idx + 1]
                    
                    if pd.notna(vuelta_num) and pd.notna(tiempo_val):
                        if str(vuelta_num).strip().isdigit():
                            registros.append({
                                'Circuito': circuito_actual,
                                'Piloto': val_header,
                                'Vuelta': int(vuelta_num),
                                'Tiempo': str(tiempo_val).strip()
                            })
                        else:
                            break
                    else:
                        break

    return pd.DataFrame(registros).drop_duplicates()

# =========================================================================
# CARGA Y PREPARACIÓN DE VARIABLES GLOBALES
# =========================================================================

opciones_fechas_combinadas = ["Campeonato Completo"]

if os.path.exists(ARCHIVO_EXCEL):
    try:
        # Carga tradicional para pestañas que usan df_datos y df_tiempos original
        xls = pd.ExcelFile(ARCHIVO_EXCEL, engine='openpyxl')
        hojas = xls.sheet_names
        
        df_datos = pd.read_excel(ARCHIVO_EXCEL, sheet_name='DATOS', engine='openpyxl') if 'DATOS' in hojas else pd.DataFrame()
        df_tiempos = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', engine='openpyxl') if 'Carga Tiempos' in hojas else pd.DataFrame()
        
        # Obtener lista de circuitos detectados para los desplegables
        if 'Carga Tiempos' in hojas:
            df_crudo_aux = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', header=None, engine='openpyxl')
            columna_a_rellenada = df_crudo_aux.iloc[:, 0].ffill()
            circuitos_detectados = []
            for c in columna_a_rellenada.dropna().astype(str).str.strip():
                c_upper = c.upper()
                if c_upper not in [x.upper() for x in circuitos_detectados] and c != "" and "NAN" not in c_upper and not c_upper.startswith("FECHA"):
                    circuitos_detectados.append(c)
            
            for idx, nombre_circuito in enumerate(circuitos_detectados):
                opciones_fechas_combinadas.append(f"Fecha {idx + 1} - {nombre_circuito}")
    except Exception as e:
        df_datos = pd.DataFrame()
        df_tiempos = pd.DataFrame()
        opciones_fechas_combinadas = ["Campeonato Completo"] + [f"Fecha {i}" for i in range(1, 11)]
else:
    df_datos = pd.DataFrame()
    df_tiempos = pd.DataFrame()
    opciones_fechas_combinadas = ["Campeonato Completo"]

# 2. MENÚ LATERAL
st.sidebar.image("https://flaticon.com", width=80) 
st.sidebar.title("Torneo _TC2000")
st.sidebar.subheader("Campeonato Interno")

opcion = st.sidebar.radio(
    "Navegación",
    ["Resumen", "Comparativa de Tiempos", "Lastre", "Duelo H2H", "Simulador de Campeonato", "Estadisticas"]
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

    if os.path.exists(ARCHIVO_EXCEL):
        try:
            # Función auxiliar interna para parsear tiempos de forma segura
            def parse_tiempo_a_segundos(val):
                if pd.isna(val):
                    return None
                try:
                    if isinstance(val, (int, float)):
                        return float(val)
                    val_str = str(val).strip().replace(',', '.')
                    if ":" in val_str:
                        partes = val_str.split(":")
                        return float(partes[0]) * 60 + float(partes[1])
                    return float(val_str)
                except:
                    return None

            # 1. CARGA DE BASE DE DATOS Y FILTRADOS DE CONTROL
            df_puntos_graf = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Tabla Final', engine='openpyxl')
            df_puntos_graf.columns = [str(c).strip() for c in df_puntos_graf.columns]
            df_puntos_graf = df_puntos_graf.dropna(subset=["PILOTO"])

            df_hoja1_graf = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')
            columna_f_graf = df_hoja1_graf.iloc[:, 5].astype(str).str.strip().str.upper().str.replace("Ó", "O", regex=False)
            
            filas_totales_fecha = columna_f_graf[columna_f_graf == "TOTAL FECHA"].index.tolist()
            filas_lastre_acum = columna_f_graf[columna_f_graf == "LASTRE ACUMULADO"].index.tolist()
            filas_c1 = columna_f_graf[columna_f_graf == "C1"].index.tolist()
            filas_c2 = columna_f_graf[columna_f_graf == "C2"].index.tolist()
            filas_pole_hoja1 = columna_f_graf[columna_f_graf == "POLE"].index.tolist()
            
            indices_pilotos_graf = {"Agus": 6, "Pablo": 8, "Juandi": 10, "Eze": 12}
            pilotos_torneo = list(indices_pilotos_graf.keys())
            
            datos_evolucion_limpios = []
            puntos_acumulados_carrera = {p: 0.0 for p in indices_pilotos_graf.keys()}
            max_puntaje_detectado = 0.0
            ultima_fecha_con_datos = 0

            for idx, fila_total in enumerate(filas_totales_fecha):
                tiene_datos_esta_fecha = False
                for piloto, col_idx in indices_pilotos_graf.items():
                    val_puntos = df_hoja1_graf.iloc[fila_total, col_idx]
                    if pd.notna(val_puntos):
                        try:
                            if float(str(val_puntos).replace(',','.',1)) > 0:
                                tiene_datos_esta_fecha = True
                        except: pass
                if tiene_datos_esta_fecha:
                    ultima_fecha_con_datos = idx + 1

            fecha_seleccionada = st.selectbox("Seleccionar Fecha o Histórico:", opciones_fechas_combinadas)

            # --- FILTRADO DINÁMICO DE PUNTOS SEGÚN LA FECHA ELEGIDA ---
            if fecha_seleccionada == "Campeonato Completo":
                df_filtrado_resumen = df_puntos_graf[["PILOTO", "PTS"]].sort_values(by="PTS", ascending=False).reset_index(drop=True)
                df_filtrado_resumen.columns = ["Piloto", "Puntos"]
                titulo_grafico = "Puntos - Campeonato Completo"
            else:
                try:
                    import re
                    # FIX: Extraemos el primer número encontrado ([0])
                    coincidencias = re.findall(r'\d+', str(fecha_seleccionada))
                    numero_fecha_detectado = int(coincidencias[0]) - 1 if coincidencias else 0
                except:
                    numero_fecha_detectado = 0
                
                fila_total_bloque_hoja1 = filas_totales_fecha[numero_fecha_detectado] if numero_fecha_detectado < len(filas_totales_fecha) else filas_totales_fecha[-1]
                tabla_puntos_fecha_individual = []
                
                for piloto_n, col_idx in indices_pilotos_graf.items():
                    val_puntos_celda = df_hoja1_graf.iloc[fila_total_bloque_hoja1, col_idx]
                    try:
                        puntos_limpios_num = float(str(val_puntos_celda).replace(",", ".").strip())
                    except:
                        puntos_limpios_num = 0.0
                    tabla_puntos_fecha_individual.append({"Piloto": piloto_n, "Puntos": puntos_limpios_num})
                
                df_filtrado_resumen = pd.DataFrame(tabla_puntos_fecha_individual).sort_values(by="Puntos", ascending=False).reset_index(drop=True)
                titulo_grafico = f"Puntos Netos - {fecha_seleccionada}"


# --- GRÁFICO DE DONA CENTRADO ---
            col_izq, col_centro, col_der = st.columns([1, 2, 1])

            with col_centro:
                st.subheader(f"🎯 Distribución de Puntos ({fecha_seleccionada})")
                
                fig_dona = px.pie(
                    df_filtrado_resumen, 
                    names="Piloto", 
                    values="Puntos",
                    hole=0.5, # Hace que sea dona en lugar de torta completa
                    template="plotly_dark",
                    color_discrete_sequence=["#e10600", "#00b0ff", "#ff9100", "#00e676"]
                )
                
                fig_dona.update_traces(
                    textinfo='label+value',
                    textposition='outside',
                    marker=dict(line=dict(color='#0f111a', width=2))
                )
                fig_dona.update_layout(
                    plot_bgcolor="#161925", paper_bgcolor="#0f111a",
                    margin=dict(l=20, r=20, t=30, b=20),
                    height=380, showlegend=False
                )
                
                st.plotly_chart(fig_dona, use_container_width=True, key="resumen_dona_fijo")
            # --- EVOLUCIÓN TEMPORAL LÍNEAS ---
            for idx, fila_total in enumerate(filas_totales_fecha[:ultima_fecha_con_datos]):
                nombre_fecha_eje_x = f"Fecha {idx + 1}"
                try:
                    circuito_real = opciones_fechas_combinadas[idx + 1].split(" - ")[1]
                except:
                    circuito_real = f"Carrera {idx + 1}"
                
                for piloto, col_idx in indices_pilotos_graf.items():
                    val_puntos_fecha = df_hoja1_graf.iloc[fila_total, col_idx]
                    puntos_fecha_limpio = str(val_puntos_fecha).replace(',', '.', 1) if pd.notna(val_puntos_fecha) else "0.0"
                    try:
                        puntos_fecha = float(puntos_fecha_limpio)
                    except:
                        puntos_fecha = 0.0
                    
                    puntos_acumulados_carrera[piloto] += puntos_fecha
                    if puntos_acumulados_carrera[piloto] > max_puntaje_detectado:
                        max_puntaje_detectado = puntos_acumulados_carrera[piloto]
                    
                    pos_c1 = "-"
                    pos_c2 = "-"
                    if idx < len(filas_c1):
                        val_c1 = df_hoja1_graf.iloc[filas_c1[idx], col_idx]
                        if pd.notna(val_c1): 
                            try: pos_c1 = f"P{int(float(str(val_c1).replace(',','.',1)))}"
                            except: pass
                    if idx < len(filas_c2):
                        val_c2 = df_hoja1_graf.iloc[filas_c2[idx], col_idx]
                        if pd.notna(val_c2): 
                            try: pos_c2 = f"P{int(float(str(val_c2).replace(',','.',1)))}"
                            except: pass
                    
                    resultado_txt = f"{pos_c1} / {pos_c2}"

                    lastre_txt = "0 Kg"
                    if idx > 0 and idx - 1 < len(filas_lastre_acum):
                        fila_lastre_anterior = filas_lastre_acum[idx - 1]
                        val_lastre = df_hoja1_graf.iloc[fila_lastre_anterior, col_idx]
                        if pd.notna(val_lastre):
                            lastre_txt = f"{str(val_lastre).upper().replace('KG', '').strip()} Kg"
                    
                    datos_evolucion_limpios.append({
                        "Piloto": piloto, "Gran Premio": nombre_fecha_eje_x, "Puntos Acumulados": puntos_acumulados_carrera[piloto],
                        "Circuito": circuito_real, "Resultado": resultado_txt, "LastreInicial": lastre_txt
                    })

            if datos_evolucion_limpios:
                st.markdown("---")
                st.subheader("📈 Evolución del Campeonato y Lastre en Vivo")
                df_melted_evolucion = pd.DataFrame(datos_evolucion_limpios)
                fig_evolucion = px.line(
                    df_melted_evolucion, x="Gran Premio", y="Puntos Acumulados", color="Piloto",
                    template="plotly_dark", markers=True, custom_data=["Circuito", "Resultado", "LastreInicial", "Piloto"]
                )
                fig_evolucion.update_traces(
                    line=dict(width=3.5), marker=dict(size=8),
                    hovertemplate="<br><b>Piloto:</b> %{customdata[3]}<br>📍 <b>Circuito:</b> %{customdata[0]}<br>🏁 <b>Resultado (C1/C2):</b> %{customdata[1]}<br>⚖️ <b>Lastre Inicial:</b> %{customdata[2]} <br>🏆 <b>Puntos Acumulados:</b> %{y} pts<extra></extra>"
                )

                fig_evolucion.update_layout(hovermode="closest", plot_bgcolor="#161925", paper_bgcolor="#0f111a", margin=dict(l=20, r=20, t=20, b=20), height=400)
                st.plotly_chart(fig_evolucion, use_container_width=True, key="evolucion_lineas_resumen_fijo")

            # =========================================================================
            # 📊 SECCIÓN DE TELEMETRÍA (RITMO MEDIO + POLES EXTRAÍDAS DE HOJA1)
            # =========================================================================
            st.markdown("---")
            st.subheader("🏎️ Telemetría y Estadísticas de Rendimiento")
            
            posiciones_c1_por_piloto = {p: [] for p in pilotos_torneo}
            posiciones_c2_por_piloto = {p: [] for p in pilotos_torneo}
            poles_totales_por_piloto = {p: 0 for p in pilotos_torneo}
            efectividad_mangas_totales = {p: [] for p in pilotos_torneo}

            # 🛠️ 1. CONTEO DE POLES DESDE HOJA1
            for f_pole in filas_pole_hoja1:
                for piloto, col_idx in indices_pilotos_graf.items():
                    columna_puntos_idx = col_idx + 1
                    if columna_puntos_idx < df_hoja1_graf.shape[1]:
                        val_pole = df_hoja1_graf.iloc[f_pole, columna_puntos_idx]
                        if pd.notna(val_pole) and str(val_pole).strip() != "":
                            try:
                                texto_pole = str(val_pole).strip().upper().replace(",0", "").replace(".0", "")
                                if texto_pole in ["1", "POLE", "1°", "🥇"]:
                                    poles_totales_por_piloto[piloto] += 1
                            except: pass

            # 🛠️ 2. MOTOR HÍBRIDO AVANZADO
            try:
                df_tiempos_aux_resumen = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', header=None, engine='openpyxl')
                df_tiempos_aux_resumen.iloc[:, 0] = df_tiempos_aux_resumen.iloc[:, 0].ffill()
                col_sesiones_resumen = df_tiempos_aux_resumen.iloc[:, 1].astype(str).str.strip().str.upper().str.replace("Ó", "O", regex=False)
                indices_pilotos_fechas_resumen = {"Agus": 2, "Pablo": 3, "Juandi": 4, "Eze": 5}
                filas_clasif_tiempos = df_tiempos_aux_resumen[col_sesiones_resumen.str.contains("CLASIF|POLE", na=False)].index.tolist()

                for idx_est in range(len(filas_c1)):
                    poleman_fecha = None
                    if idx_est < len(filas_pole_hoja1):
                        f_pole_actual = filas_pole_hoja1[idx_est]
                        for piloto, col_idx in indices_pilotos_graf.items():
                            if col_idx + 1 < df_hoja1_graf.shape[1]:
                                val_p_celda = df_hoja1_graf.iloc[f_pole_actual, col_idx + 1]
                                if pd.notna(val_p_celda) and str(val_p_celda).strip() in ["1", "1.0"]:
                                    poleman_fecha = piloto
                                    break

                    grilla_c1_largada = {p: 4.0 for p in pilotos_torneo}
                    
                    if idx_est < len(filas_clasif_tiempos):
                        f_qualy = filas_clasif_tiempos[idx_est]
                        tiempos_fecha = {}
                        for p, col_idx in indices_pilotos_fechas_resumen.items():
                            seg = parse_tiempo_a_segundos(df_tiempos_aux_resumen.iloc[f_qualy, col_idx])
                            if seg is not None and seg > 30.0 and p != poleman_fecha:
                                tiempos_fecha[p] = seg
                        
                        if poleman_fecha:
                            grilla_c1_largada[poleman_fecha] = 1.0
                        
                        if tiempos_fecha:
                            pilotos_restantes_ordenados = sorted(tiempos_fecha, key=tiempos_fecha.get)
                            puesto_disponible = 2.0
                            for p_name in pilotos_restantes_ordenados:
                                grilla_c1_largada[p_name] = puesto_disponible
                                puesto_disponible += 1.0

                    for piloto, col_idx in indices_pilotos_graf.items():
                        val_c1 = df_hoja1_graf.iloc[filas_c1[idx_est], col_idx]
                        if pd.notna(val_c1):
                            try:
                                llegada_c1 = float(str(val_c1).upper().replace("P", "").strip())
                                if llegada_c1 > 0:
                                    posiciones_c1_por_piloto[piloto].append(llegada_c1)
                                    largada_c1 = grilla_c1_largada[piloto]
                                    
                                    if largada_c1 == llegada_c1:
                                        efec_c1 = 100.0
                                    elif largada_c1 > llegada_c1:
                                        efec_c1 = ((largada_c1 - llegada_c1) / (largada_c1 - 1.0)) * 100
                                    else:
                                        efec_c1 = (((4.0 - largada_c1) - (llegada_c1 - largada_c1)) / (4.0 - largada_c1)) * 100
                                    
                                    efectividad_mangas_totales[piloto].append(efec_c1)
                            except: pass

                    if idx_est < len(filas_c2):
                        for piloto, col_idx in indices_pilotos_graf.items():
                            val_c2 = df_hoja1_graf.iloc[filas_c2[idx_est], col_idx]
                            if pd.notna(val_c2):
                                try:
                                    llegada_c2 = float(str(val_c2).upper().replace("P", "").strip())
                                    if llegada_c2 > 0:
                                        posiciones_c2_por_piloto[piloto].append(llegada_c2)
                                        largada_c2 = 5.0 - grilla_c1_largada[piloto]
                                        
                                        if largada_c2 == llegada_c2:
                                            efec_c2 = 100.0
                                        elif largada_c2 > llegada_c2:
                                            efec_c2 = ((largada_c2 - llegada_c2) / (largada_c2 - 1.0)) * 100
                                        else:
                                            efec_c2 = (((4.0 - largada_c2) - (llegada_c2 - largada_c2)) / (4.0 - largada_c2)) * 100
                                        
                                        efectividad_mangas_totales[piloto].append(efec_c2)
                                except: pass
            except Exception as e_proc:
                st.warning(f"Aviso en cálculo de grillas: {e_proc}")

            # --- COMPILACIÓN DEL REPORTE FINAL ---
            reporte_tarjetas = []
            for piloto in pilotos_torneo:
                lista_c1 = posiciones_c1_por_piloto.get(piloto, [])
                lista_c2 = posiciones_c2_por_piloto.get(piloto, [])
                prom_c1 = sum(lista_c1) / len(lista_c1) if lista_c1 else 0.0
                prom_c2 = sum(lista_c2) / len(lista_c2) if lista_c2 else 0.0
                poles_reales = poles_totales_por_piloto.get(piloto, 0)
                
                lista_efec = efectividad_mangas_totales.get(piloto, [])
                prom_efec_carrera = sum(lista_efec) / len(lista_efec) if lista_efec else 0.0
                
                gps_disputados_c1 = len(lista_c1)
                porcentaje_efectividad_pole = (poles_reales / gps_disputados_c1) * 100 if gps_disputados_c1 > 0 else 0.0
                
                total_mangas = lista_c1 + lista_c2
                prom_general = sum(total_mangas) / len(total_mangas) if total_mangas else 0.0

                reporte_tarjetas.append({
                    "Piloto": piloto, "Promedio C1": prom_c1, "Promedio C2": prom_c2, 
                    "Promedio General": prom_general, "Poles": poles_reales,
                    "Efectividad_Pole": porcentaje_efectividad_pole,
                    "Efectividad_Carrera": prom_efec_carrera,
                    "GPs": gps_disputados_c1
                })

            df_reporte_tarjetas = pd.DataFrame(reporte_tarjetas).sort_values(by="Promedio General")

            # --- RENDERIZADO DE LAS TARJETAS (ESTILO LASTRE) ---
            st.markdown("""
            <style>
            .tarjeta-simulacion-f1 {
                background-color: #161925; border-left: 5px solid #e10600; border-radius: 10px;
                padding: 18px; margin-bottom: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.5);
            }
            .titulo-simulacion-f1 { margin: 0; font-size: 11px; color: #9fa6b2; text-transform: uppercase; font-weight: bold; letter-spacing: 1px; }
            .piloto-simulacion-f1 { margin: 0; padding: 4px 0; color: #ffffff; font-size: 26px; font-weight: bold; }
            .divisor-simulacion-f1 { margin: 8px 0; border-color: #2a2f45; }
            .bloque-valores-f1 { display: flex; justify-content: space-between; align-items: center; }
            .sub-metrica-f1 { text-align: left; }
            .sub-metrica-der-f1 { text-align: right; }
            .texto-gris-f1 { margin: 0; font-size: 12px; color: #9fa6b2; }
            .texto-blanco-bold-f1 { margin: 0; font-size: 20px; color: #ffffff; font-weight: bold; font-family: monospace; }
            .texto-rojo-bold-f1 { margin: 0; font-size: 15px; color: #e10600; font-weight: bold; font-family: monospace; }
            </style>
            """, unsafe_allow_html=True)

            cols_grid = st.columns(4)
            for idx_c, row in df_reporte_tarjetas.reset_index(drop=True).iterrows():
                p_name = row["Piloto"]
                p_gen = f"P{row['Promedio General']:.1f}" if row["Promedio General"] > 0 else "-"
                p_c1 = f"P{row['Promedio C1']:.1f}" if row["Promedio C1"] > 0 else "-"
                p_c2 = f"P{row['Promedio C2']:.1f}" if row["Promedio C2"] > 0 else "-"
                p_poles = int(row["Poles"])
                p_efec_pole = row["Efectividad_Pole"]
                p_efec_carrera = row["Efectividad_Carrera"]
                p_gps = int(row["GPs"])
                
                color_borde_pista = "#00e676" if idx_c == 0 else "#e10600"
                
                with cols_grid[idx_c % 4]:
                    st.markdown(f"""
                        <div class="tarjeta-simulacion-f1" style="border-left: 5px solid {color_borde_pista};">
                            <p class="titulo-simulacion-f1">Rendimiento en Pista ({p_gps} GPs)</p>
                            <h3 class="piloto-simulacion-f1">{p_name}</h3>
                            <hr class="divisor-simulacion-f1">
                            <div class="bloque-valores-f1">
                                <div class="sub-metrica-f1">
                                    <p class="texto-gris-f1">Ritmo General:</p>
                                    <p class="texto-blanco-bold-f1" style="color: #00e676;">{p_gen}</p>
                                    <p class="texto-gris-f1" style="font-size:11px; margin-top:4px; color: #ff9100; font-weight:bold;">🏁 Efec. Race: {p_efec_carrera:.1f}%</p>
                                </div>
                                <div class="sub-metrica-der-f1">
                                    <p class="texto-gris-f1">Poles (Sábado):</p>
                                    <p class="texto-rojo-bold-f1" style="color: #00b0ff; font-size:22px;">🥇 {p_poles}</p>
                                    <p class="texto-gris-f1" style="font-size:11px; margin-top:4px; color: #9fa6b2;">⚡ Efec. Pole: {p_efec_pole:.1f}%</p>
                                </div>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)

            st.write("#### 📊 Tabla de Resumen Detallada")
            st.dataframe(
                df_reporte_tarjetas.assign(
                    **{
                        "Promedio General": lambda x: x["Promedio General"].apply(lambda v: f"P{v:.1f}" if v > 0 else "-"),
                        "Promedio C1": lambda x: x["Promedio C1"].apply(lambda v: f"P{v:.1f}" if v > 0 else "-"),
                        "Promedio C2": lambda x: x["Promedio C2"].apply(lambda v: f"P{v:.1f}" if v > 0 else "-"),
                        "Poles": lambda x: x["Poles"].apply(lambda v: f"🥇 {int(v)} Poles"),
                        "Efectividad_Pole": lambda x: x["Efectividad_Pole"].apply(lambda v: f"{v:.1f}%"),
                        "Efectividad_Carrera": lambda x: x["Efectividad_Carrera"].apply(lambda v: f"{v:.1f}%")
                    }
                ),
                use_container_width=True, hide_index=True
            )

            rey_pole = max(poles_totales_por_piloto, key=poles_totales_por_piloto.get)
            max_poles = poles_totales_por_piloto[rey_pole]
            if max_poles > 0:
                st.success(f"⏱️ **Rey de los Sábados:** El piloto con más Pole Positions es **{rey_pole}** con un total de **{max_poles} Poles**.")
                st.caption("ℹ️ **Nota sobre la Efectividad de Carrera:** Este porcentaje mide el aprovechamiento de posiciones ganadas desde la posición inicial de largada en cada manga.")

        except Exception as e:
            st.error(f"Error general en la pestaña Resumen: {e}")



elif opcion == "Comparativa de Tiempos":
    st.title("⏱️ Diferencia de Ritmo y Poles (Histórico vs Fecha)")
    st.write("Análisis estadístico milimétrico basado en tu pestaña 'Carga Tiempos'")

    if df_tiempos is not None:
        fecha_tiempos_sel_combinada = st.selectbox("Seleccionar Período a Analizar:", opciones_fechas_combinadas, key="selector_tiempos_global")
        
        if fecha_tiempos_sel_combinada == "Campeonato Completo":
            fecha_tiempos_sel = "Campeonato Completo"
        else:
            fecha_tiempos_sel = str(fecha_tiempos_sel_combinada).split(" - ")[0]
        
        df_tiempos_crudo = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Carga Tiempos', header=None, engine='openpyxl')
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

        # 🛠️ NUEVA FUNCIÓN IN SITU PARA PASAR SEGUNDOS A "MM:SS.mmm"
        def formatear_segundos_a_vuelta(segundos_totales):
            if segundos_totales is None or pd.isna(segundos_totales) or segundos_totales == 999.0:
                return "-"
            try:
                minutos = int(segundos_totales // 60)
                segundos_puros = segundos_totales % 60
                return f"{minutos:02d}:{segundos_puros:06.3f}".replace(".", ",")
            except:
                return f"{segundos_totales:.3f}s"

        css_tiempos = """<style>.columna-tiempos-box { background-color: #0f111a; border: 1px solid #2a2f45; border-radius: 12px; padding: 15px; margin-bottom: 20px; } .tarjeta-tiempo-f1 { background-color: #161925; border-left: 5px solid #e10600; border-radius: 8px; padding: 10px 14px; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 8px rgba(0,0,0,0.3); } .info-izq-tiempo { display: flex; flex-direction: column; } .piloto-tiempo-f1 { color: #ffffff; font-size: 15px; font-weight: bold; } .tiempo-vuelta-f1 { color: #00ff88; font-size: 13px; font-family: monospace; margin-top: 1px; } .gap-der-tiempo { text-align: right; font-size: 14px; font-weight: bold; font-family: monospace; }</style>"""
        st.markdown(css_tiempos, unsafe_allow_html=True)


        # =========================================================================
        # 📈 DETECTOR COMPLETO PARA EL CAMPEONATO COMPLETO
        # =========================================================================
        if fecha_tiempos_sel == "Campeonato Completo":
            st.header("📈 Resumen Global Acumulado (Todas las Carreras Disputadas)")
            
            def calcular_promedios_historicos_python(palabra_sesion, titulo_seccion, key_graf):
                tiempos_acumulados = {p: [] for p in pilotos}
                
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
                        if p in promedios_finales and promedios_finales[p] is not None:
                            brecha = promedios_finales[p] - tiempo_base
                            t_formateado = formatear_segundos_a_vuelta(promedios_finales[p])
                            b_formateada = "Líder" if brecha == 0 else f"+{brecha:.3f}s".replace(".", ",")
                            tabla_global.append({"Piloto": p, "Tiempo": t_formateado, "Brecha con Líder": b_formateada, "Orden_Num": brecha})
                        else:
                            tabla_global.append({"Piloto": p, "Tiempo": "-", "Brecha con Líder": "-", "Orden_Num": 999.0})
                    
                    df_res = pd.DataFrame(tabla_global).sort_values("Orden_Num").reset_index(drop=True)
                    
                    # --- COMPONENTE VISUAL EN TARJETAS HORIZONTALES ---
                    st.markdown(f'<div class="columna-tiempos-box"><span style="color: #ffffff; font-size: 16px; font-weight: bold; display: block; margin-bottom: 12px;">📋 {titulo_seccion}</span>', unsafe_allow_html=True)

                    for idx_t, row_t in df_res.iterrows():
                        color_borde = "#00e676" if idx_t == 0 else "#e10600"
                        color_texto_gap = "#00e676" if row_t["Brecha con Líder"] == "Líder" else "#ff9100"
                        
                        st.markdown(f"""
                            <div class="tarjeta-tiempo-f1" style="border-left: 5px solid {color_borde};">
                                <div class="info-izq-tiempo">
                                    <span class="piloto-tiempo-f1">#{idx_t + 1} — {row_t['Piloto']}</span>
                                    <span class="tiempo-vuelta-f1">⏱️ {row_t['Tiempo']}</span>
                                </div>
                                <div class="gap-der-tiempo" style="color: {color_texto_gap};">
                                    {row_t['Brecha con Líder']}
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)

            col1, col2, col3 = st.columns(3)
            with col1: calcular_promedios_historicos_python("CLASIF", "Clasificación (Poles)", "c_hist")
            with col2: calcular_promedios_historicos_python("CARRERA 1", "Carrera 1", "c1_hist")
            with col3: calcular_promedios_historicos_python("CARRERA 2", "Carrera 2", "c2_hist")

        # =========================================================================
        # 📅 DETECTOR PARA FECHA INDIVIDUAL SELECCIONADA
        # =========================================================================
        else:
            try:
                idx_fecha = int(fecha_tiempos_sel.replace("Fecha ", "")) - 1
                circuito_buscado = circuitos_ordenados[idx_fecha]
            except (IndexError, ValueError):
                circuito_buscado = None

            if circuito_buscado:
                st.header(f"📍 Tiempos en {circuito_buscado} ({fecha_tiempos_sel_combinada})")
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
                                t_formateado = formatear_segundos_a_vuelta(tiempos_pilotos_fecha[p])
                                b_formateada = "Líder" if brecha == 0 else f"+{brecha:.3f}s".replace(".", ",")
                                tabla_brechas.append({"Piloto": p, "Tiempo": t_formateado, "Brecha con Líder": b_formateada, "Orden_Num": brecha})
                            else:
                                tabla_brechas.append({"Piloto": p, "Tiempo": "-", "Brecha con Líder": "-", "Orden_Num": 999.0})
                        
                        df_res_ind = pd.DataFrame(tabla_brechas).sort_values(by="Orden_Num").reset_index(drop=True)
                        
                        # --- COMPONENTE VISUAL INDIVIDUAL EN TARJETAS HORIZONTALES ---
                        st.markdown(f'<div class="columna-tiempos-box">📋 <b>{tipo_sesion}</b>', unsafe_allow_html=True)
                        for idx_t, row_t in df_res_ind.iterrows():
                            color_borde = "#00e676" if idx_t == 0 else "#e10600"
                            color_texto_gap = "#00e676" if row_t["Brecha con Líder"] == "Líder" else "#ff9100"
                            
                            st.markdown(f"""
                                <div class="tarjeta-tiempo-f1" style="border-left: 5px solid {color_borde};">
                                    <div class="info-izq-tiempo">
                                        <span class="piloto-tiempo-f1">#{idx_t + 1} — {row_t['Piloto']}</span>
                                        <span class="tiempo-vuelta-f1">⏱️ {row_t['Tiempo']}</span>
                                    </div>
                                    <div class="gap-der-tiempo" style="color: {color_texto_gap};">
                                        {row_t['Brecha con Líder']}
                                    </div>
                                </div>
                            """, unsafe_allow_html=True)
                        st.markdown('</div>', unsafe_allow_html=True)

                col1, col2, col3 = st.columns(3)
                with col1: analizar_sesion_individual_dinamica(fila_clasif, "Clasificación")
                with col2: analizar_sesion_individual_dinamica(fila_c1, "Carrera 1")
                with col3: analizar_sesion_individual_dinamica(fila_c2, "Carrera 2")
            else:
                st.warning(f"La {fecha_tiempos_sel} aún no tiene datos cargados en el Excel.")


elif opcion == "Lastre":
    st.title("⚖️ Sistema de Lastre Técnico")
    st.write("Consulta el peso acumulado en pista.")

    if os.path.exists(ARCHIVO_EXCEL):
        try:
            # 1. LEER LAS PLANILLAS DE FORMA SEGURA
            df_hoja1_sim = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')
            columna_f_sim = df_hoja1_sim.iloc[:, 5].astype(str).str.strip().str.upper()
            
            df_puntos_aux = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Tabla Final', engine='openpyxl')
            df_puntos_aux.columns = [str(c).strip() for c in df_puntos_aux.columns]

            # Mapeo de columnas oficiales de pilotos en tu Hoja1
            indices_pilotos_hoja1 = {"Agus": 6, "Pablo": 8, "Juandi": 10, "Eze": 12}
            pilotos_torneo = list(indices_pilotos_hoja1.keys())

            # Detectamos cuántas fechas se corrieron realmente leyendo la Tabla Final de puntos
            columnas_fechas_reales = [col for col in df_puntos_aux.columns if str(col).strip().upper().startswith("FECHA")]
            ultima_fecha_real_num = 0
            for idx_f, col_f in enumerate(columnas_fechas_reales):
                if df_puntos_aux[col_f].astype(float).sum() > 0:
                    ultima_fecha_real_num = idx_f + 1
            if ultima_fecha_real_num == 0:
                ultima_fecha_real_num = 7

            filas_lastre_fecha_real = columna_f_sim[columna_f_sim.str.contains("LASTRE FECHA|LASTRE.*FECHA", na=False)].index.tolist()
            filas_lastre_acum_real = columna_f_sim[columna_f_sim.str.contains("LASTRE ACUMULADO|LASTRE.*ACUM", na=False)].index.tolist()

            if len(filas_lastre_fecha_real) > 0:
                total_fechas_detectadas = len(filas_lastre_fecha_real)
                opciones_lastre_dinamicas = [opt for opt in opciones_fechas_combinadas if opt != "Campeonato Completo"][:total_fechas_detectadas]
                fecha_sel = st.selectbox("Seleccionar Fecha para Consultar Lastre Técnico:", opciones_lastre_dinamicas)
                
                import re
                numeros_encontrados = re.findall(r'\d+', str(fecha_sel))
                idx_fecha_sel = int(numeros_encontrados[0]) - 1 if numeros_encontrados else 0

                # Armamos la tabla para renderizar los kilos en pista actuales
                tabla_lastre = []
                fila_pista = filas_lastre_acum_real[idx_fecha_sel] if idx_fecha_sel < len(filas_lastre_acum_real) else None
                fila_generado = filas_lastre_fecha_real[idx_fecha_sel] if idx_fecha_sel < len(filas_lastre_fecha_real) else None
                
                for p in pilotos_torneo:
                    col_idx = indices_pilotos_hoja1[p]
                    k_pista = df_hoja1_sim.iloc[fila_pista, col_idx] if fila_pista else 0
                    k_gen = df_hoja1_sim.iloc[fila_generado, col_idx] if fila_generado else 0
                    tabla_lastre.append({
                        "Piloto": p,
                        "Lastre en Pista (kg)": float(str(k_pista).upper().replace("KG","").strip()) if pd.notna(k_pista) else 0.0,
                        "Lastre Generado (kg)": float(str(k_gen).upper().replace("KG","").strip()) if pd.notna(k_gen) else 0.0
                    })
                
                df_lastre_render = pd.DataFrame(tabla_lastre)
                
                # Dibujamos las tarjetas individuales arriba
                cols_cards = st.columns(len(pilotos_torneo))
                for i, row in enumerate(df_lastre_render.iterrows()):
                    with cols_cards[i % 4]:
                        st.metric(label=f"Auto de {row[1]['Piloto']}", value=f"{row[1]['Lastre en Pista (kg)']} Kg", delta=f"+{row[1]['Lastre Generado (kg)']} Kg ganados")
            else:
                st.warning("No se detectaron las filas de 'Lastre Fecha' en la 'Hoja1'.")

        except Exception as e:
            st.error(f"Error al calcular el lastre desde el Excel: {e}")
    else:
        st.warning(f"No se encontró el archivo '{ARCHIVO_EXCEL}'.")
elif opcion == "Posiciones":
    if os.path.exists(ARCHIVO_EXCEL):
        # 1. CARGA DIRECTA Y FRESCA DE LA HOJA1 EN CADA MOVIMIENTO DEL SELECTOR
        df_hoja1 = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')
        
        # # USAMOS LA TABLA GLOBAL QUE NUNCA FALLA Y YA TIENE LAS FECHAS COMPUTADAS
        # Buscamos en qué posición (0 para Fecha 1, 7 para Fecha 8) está el circuito que eligieron
        idx_fecha_sel = opciones_fechas_combinadas.index(fecha_seleccionada) if fecha_seleccionada in opciones_fechas_combinadas else 0

        tabla_fecha_pura = []
        tabla_campeonato_historico = []

        # Recorremos los 4 pilotos de tu torneo de Assetto Corsa
        for piloto in ["Agus", "Pablo", "Juandi", "Eze"]:
            # Filtramos el historial ordenado de ese piloto en el DataFrame del gráfico
            datos_piloto = df_melted_evolucion[df_melted_evolucion['Piloto'] == piloto].reset_index(drop=True)
            
            puntos_fecha_puros = 0.0
            puntos_acumulados_historicos = 0.0
            
            if idx_fecha_sel < len(datos_piloto):
                # Extraemos el total acumulado real que llevaba hasta esa fecha exacta
                puntos_acumulados_historicos = float(datos_piloto.iloc[idx_fecha_sel]['Puntos Acumulados'])
                
                if idx_fecha_sel == 0:
                    # En la Fecha 1, los puntos de la carrera son iguales al acumulado
                    puntos_fecha_puros = puntos_acumulados_historicos
                else:
                    # Aplicamos tu fórmula: Acumulado de hoy MENOS acumulado de la fecha anterior
                    puntos_acum_anterior = float(datos_piloto.iloc[idx_fecha_sel - 1]['Puntos Acumulados'])
                    puntos_fecha_puros = puntos_acumulados_historicos - puntos_acum_anterior

            tabla_fecha_pura.append({"Piloto": piloto, "Puntos de la Fecha": puntos_fecha_puros})
            tabla_campeonato_historico.append({"Piloto": piloto, "Puntos Totales": puntos_acumulados_historicos})

        # 5. CREACIÓN Y ORDENAMIENTO DE LAS TABLAS (Conecta directo con tu código visual de abajo)
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
            

    
    else:
        st.warning(f"No se encontró el archivo '{ARCHIVO_EXCEL}'.")


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
elif opcion == "Simulador de Campeonato":
    st.title("🔮 Simulador y Proyecciones del Campeonato")
    st.write("Calculadora matemática de título multi-fecha y evolución del lastre técnico en vivo.")

    # 1. CARGA DE DATOS SEGURA DESDE AMBAS PESTAÑAS
    try:
        df_puntos_sim = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Tabla Final', engine='openpyxl')
        df_puntos_sim.columns = [str(c).strip() for c in df_puntos_sim.columns]
        
        df_hoja1_sim = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')
        columna_f_sim = df_hoja1_sim.iloc[:, 5].astype(str).str.strip().str.upper()
    except Exception as e:
        st.error(f"❌ No se pudo cargar la base de datos para el simulador: {e}")
        df_puntos_sim = None

    if df_puntos_sim is not None and df_hoja1_sim is not None:
        df_puntos_sim = df_puntos_sim.dropna(subset=["PILOTO"])
        pilotos_torneo = sorted(df_puntos_sim["PILOTO"].unique())
        indices_pilotos_hoja1 = {"Agus": 6, "Pablo": 8, "Juandi": 10, "Eze": 12}

        # ---------------------------------------------------------------------
        # 📊 DETECCIÓN REAL DE FECHAS DISPUTADAS BASADO EN PUNTOS
        # ---------------------------------------------------------------------
        columnas_fechas_reales = [col for col in df_puntos_sim.columns if str(col).strip().upper().startswith("FECHA")]
        ultima_fecha_real_num = 0
        
        for idx, col_f in enumerate(columnas_fechas_reales):
            if df_puntos_sim[col_f].astype(float).sum() > 0:
                ultima_fecha_real_num = idx + 1

        if ultima_fecha_real_num == 0:
            ultima_fecha_real_num = 8

        filas_lastre_fecha_real = columna_f_sim[columna_f_sim.str.contains("LASTRE FECHA|LASTRE.*FECHA", na=False)].index.tolist()
        filas_lastre_acum_real = columna_f_sim[columna_f_sim.str.contains("LASTRE ACUMULADO|LASTRE.*ACUM", na=False)].index.tolist()
        
        historial_generado_por_piloto = {p: [] for p in pilotos_torneo}
        lastre_inicial_proyeccion = {p: 0.0 for p in pilotos_torneo}
        puntos_reales_actuales = {p: 0.0 for p in pilotos_torneo}

        for p in pilotos_torneo:
            fila_p = df_puntos_sim[df_puntos_sim["PILOTO"] == p]
            # 🛠️ CORRECCIÓN QUIRÚRGICA: Agregamos [0] para extraer el escalar numérico de la matriz de NumPy
            puntos_reales_actuales[p] = float(fila_p["PTS"].values[0]) if not fila_p.empty else 0.0
            
            col_idx = indices_pilotos_hoja1.get(p, 6)
            
            for i in range(min(ultima_fecha_real_num, len(filas_lastre_fecha_real))):
                f_idx = filas_lastre_fecha_real[i]
                val = df_hoja1_sim.iloc[f_idx, col_idx]
                peso = float(str(val).upper().replace("KG","").strip()) if pd.notna(val) and str(val).strip() != "" else 0.0
                historial_generado_por_piloto[p].append(peso)
            
            if len(filas_lastre_acum_real) >= ultima_fecha_real_num:
                f_acum_idx = filas_lastre_acum_real[ultima_fecha_real_num - 1]
                val_ultimo_acum = df_hoja1_sim.iloc[f_acum_idx, col_idx]
                lastre_inicial_proyeccion[p] = float(str(val_ultimo_acum).upper().replace("KG","").strip()) if pd.notna(val_ultimo_acum) and str(val_ultimo_acum).strip() != "" else 0.0

        # ---------------------------------------------------------------------
        # 🎛️ CONTROLES INTERACTIVOS DE PROYECCIÓN MULTI-FECHA
        # ---------------------------------------------------------------------
        st.markdown("---")
        st.subheader("🏁 Configuración de la Proyección Futura")
        st.write(f"📢 Campeonato actual procesado hasta la **Fecha {ultima_fecha_real_num}**. Podés simular las fechas restantes:")
        
        fechas_restantes_max = max(1, 10 - ultima_fecha_real_num)
        cant_fechas_a_proyectar = st.number_input(f"¿Cuántas fechas querés proyectar hacia adelante? (Quedan {fechas_restantes_max} para llegar a la 10)", min_value=1, max_value=int(fechas_restantes_max), value=1, step=1)
        
        escala_c1 = {1: 25, 2: 18, 3: 15, 4: 12}
        escala_c2 = {pos: pts * 0.75 for pos, pts in escala_c1.items()}

        escala_generado_lastre = {1: 40.0, 2: 20.0, 3: 10.0, 4: 0.0}

        puntos_simulados_acum = puntos_reales_actuales.copy()
        lastre_actual_simulado = lastre_inicial_proyeccion.copy()
        historial_generado_simulado = {p: list(historial_generado_por_piloto[p]) for p in pilotos_torneo}

        for f_futura in range(cant_fechas_a_proyectar):
            num_fecha_actual_sim = ultima_fecha_real_num + f_futura + 1
            st.markdown(f"### 📅 Simular Fecha {num_fecha_actual_sim}")
            
            st.write(f"**⚖️ Telemetría Inicial para la Fecha {num_fecha_actual_sim}:**")
            cols_kilos = st.columns(len(pilotos_torneo))
            for i, p in enumerate(pilotos_torneo):
                with cols_kilos[i]:
                    pts_actuales_piloto = puntos_simulados_acum[p]
                    st.metric(
                        label=f"Auto de {p}", 
                        value=f"{lastre_actual_simulado[p]:.1f} Kg",
                        delta=f"{pts_actuales_piloto:.2f} pts iniciales",
                        delta_color="off"
                    )

            col_c1, col_c2 = st.columns(2)
            posiciones_f_c1 = {}
            posiciones_f_c2 = {}

            with col_c1:
                st.markdown(f"🔺 **Carrera 1 - Fecha {num_fecha_actual_sim}**")
                for piloto in pilotos_torneo:
                    pos_c1 = st.selectbox(
                        f"{piloto} - Posición C1:", options=[1, 2, 3, 4], 
                        index=pilotos_torneo.index(piloto) % 4, 
                        key=f"c1_f{num_fecha_actual_sim}_{piloto}"
                    )
                    posiciones_f_c1[piloto] = pos_c1

            with col_c2:
                st.markdown(f"🔻 **Carrera 2 - Fecha {num_fecha_actual_sim}**")
                for piloto in pilotos_torneo:
                    pos_c2 = st.selectbox(
                        f"{piloto} - Posición C2:", options=[1, 2, 3, 4], 
                        index=(pilotos_torneo.index(piloto) + 1) % 4, 
                        key=f"c2_f{num_fecha_actual_sim}_{piloto}"
                    )
                    posiciones_f_c2[piloto] = pos_c2

            puntos_ganados_esta_fecha = {}
            for piloto in pilotos_torneo:
                pts_f = escala_c1.get(posiciones_f_c1[piloto], 0) + escala_c2.get(posiciones_f_c2[piloto], 0)
                puntos_ganados_esta_fecha[piloto] = pts_f
                puntos_simulados_acum[piloto] += pts_f

            pilotos_ordenados_fecha = sorted(puntos_ganados_esta_fecha, key=puntos_ganados_esta_fecha.get, reverse=True)
            
            for rango_pos, piloto in enumerate(pilotos_ordenados_fecha):
                posicion_final_fecha = rango_pos + 1
                kilos_generados_hoy = escala_generado_lastre.get(posicion_final_fecha, 0.0)
                
                historial_generado_simulado[piloto].append(kilos_generados_hoy)
                
                idx_actual_historial = len(historial_generado_simulado[piloto]) - 1
                idx_a_restar = idx_actual_historial - 2
                
                kilos_a_restar_tres_atras = 0.0
                if idx_a_restar >= 0 and idx_a_restar < len(historial_generado_simulado[piloto]):
                    kilos_a_restar_tres_atras = historial_generado_simulado[piloto][idx_a_restar]
                
                nuevo_lastre_calculado = lastre_actual_simulado[piloto] + kilos_generados_hoy - kilos_a_restar_tres_atras
                lastre_actual_simulado[piloto] = max(0.0, nuevo_lastre_calculado)

            st.markdown("---")

        # ---------------------------------------------------------------------
        # 📊 RENDIMIENTO Y TABLA FINAL DE LA PROYECCIÓN COMPLETA
        # ---------------------------------------------------------------------
        st.subheader(f"🏆 Resultado de la Proyección General (Luego de {cant_fechas_a_proyectar} Fechas Simuladas)")
        
        tabla_proyeccion_final = []
        for piloto in pilotos_torneo:
            pts_reales_iniciales = puntos_reales_actuales[piloto]
            pts_totales_proyectados = puntos_simulados_acum[piloto]
            
            tabla_proyeccion_final.append({
                "Piloto": piloto,
                "Pts Sumados Simulación": pts_totales_proyectados - pts_reales_iniciales,
                "Puntos Finales Proyectados": pts_totales_proyectados,
                "Lastre Próxima Fecha": lastre_actual_simulado[piloto]
            })

        df_proyeccion_final = pd.DataFrame(tabla_proyeccion_final).sort_values(by="Puntos Finales Proyectados", ascending=False)
        df_proyeccion_final = df_proyeccion_final.reset_index(drop=True)

        lider_real = max(puntos_reales_actuales, key=puntos_reales_actuales.get)
        lider_proyectado = df_proyeccion_final.loc[0, "Piloto"]
        puntos_lider_proy = df_proyeccion_final.loc[0, "Puntos Finales Proyectados"]

        if lider_real != lider_proyectado:
            segundo_proy_puntos = df_proyeccion_final.loc[1, "Puntos Finales Proyectados"] if len(df_proyeccion_final) > 1 else puntos_lider_proy
            st.error(f"🔥 ¡CAMBIO DE CORONA EN LA PROYECCIÓN! Al cabo de las fechas simuladas, **{lider_proyectado}** se consagraría líder del campeonato superando al escolta por {(puntos_lider_proy - segundo_proy_puntos):.2f} pts.")
        else:
            st.success(f"👑 ¡Automático! **{lider_proyectado}** resiste la presión y mantiene la punta del campeonato con un total proyectado de {puntos_lider_proy:.2f} pts.")
    # =========================================================================
    # 🎯 ANÁLISIS DE DEFINICIÓN DE CAMPEONATO (ÚLTIMA FECHA)
    # =========================================================================
    fecha_final_simulada_num = ultima_fecha_real_num + cant_fechas_a_proyectar
    fechas_restantes_campeonato = max(0, 10 - fecha_final_simulada_num)

    if fechas_restantes_campeonato == 1:
        st.markdown("---")
        st.subheader("🎯 Análisis de Definición de Campeonato")

        df_ranking_actual = df_proyeccion_final.sort_values(by="Puntos Finales Proyectados", ascending=False).reset_index(drop=True)
        lider_actual_escenario = df_ranking_actual.loc[0, "Piloto"]
        pts_lider = df_ranking_actual.loc[0, "Puntos Finales Proyectados"]
        
        pts_segundo_lugar = df_ranking_actual.loc[1, "Puntos Finales Proyectados"] if len(df_ranking_actual) > 1 else pts_lider
        brecha_lider = pts_lider - pts_segundo_lugar

        # 1. Selector de tipo de puntaje único (antes del cartel azul)
        st.markdown("<h4 style='color: #262730; margin-bottom: -10px;'>🏁 Selecciona el tipo de puntaje para la última fecha:</h4>", unsafe_allow_html=True)
        
        tipo_fecha = st.radio(
            "Selecciona el tipo de puntaje para la última fecha:",
            ["Normal (Máx. 46.75 pts)", "Especial x1.5 (Máx. 70.125 pts)"],
            horizontal=True,
            label_visibility="collapsed",
            key="tipo_puntaje_decisivo_final"
        )
        
        if "Especial" in tipo_fecha:
            puntos_maximos_por_fecha = 46.75 * 1.5  
        else:
            puntos_maximos_por_fecha = 46.75

        limite_seguro_lider = max(21.0, puntos_maximos_por_fecha - brecha_lider)
        puntos_ejemplo_lider = max(21.0, limite_seguro_lider - 0.25)

        # 2. Evaluación previa unificada y estricta para el resumen superior
        siguen_pelea = []
        sin_chances = []
        
        for _, row_esc in df_ranking_actual.iterrows():
            p_name = row_esc["Piloto"]
            p_pts = row_esc["Puntos Finales Proyectados"]
            if p_name == lider_actual_escenario:
                continue
            p_brecha = pts_lider - p_pts
            puntos_necesarios_perseguidor = p_brecha + puntos_ejemplo_lider
            
            if p_brecha > puntos_maximos_por_fecha or puntos_necesarios_perseguidor > puntos_maximos_por_fecha:
                sin_chances.append(p_name)
            else:
                siguen_pelea.append((p_name, p_brecha))

        # 3. Resumen visual superior sincronizado
        cols_resumen = st.columns(2)
        with cols_resumen[0]:
            st.write("🟢 **Siguen en la Pelea Matemática:**")
            st.write(f"👑 {lider_actual_escenario} (Líder actual)")
            for p_n, p_b in siguen_pelea:
                st.write(f"🏎️ {p_n} (A {p_b:.2f} pts del líder)")
        with cols_resumen[1]:
            st.write("🔴 **Matemáticamente Sin Chances:**")
            if sin_chances:
                for p_n in sin_chances:
                    st.write(f"❌ {p_n}")
            else:
                st.write("Ninguno")

        st.markdown("<br>", unsafe_allow_html=True)

        # 4. Cartel azul informativo justo debajo del selector
        st.info(f"ℹ️ Fecha decisiva ({tipo_fecha}). Líder actual: **{lider_actual_escenario}** ({pts_lider:.2f} pts). Máximo a repartir: **{puntos_maximos_por_fecha:.3f} pts**.")

        # 5. Expanders de detalle por piloto
        for _, row_esc in df_ranking_actual.iterrows():
            piloto_name = row_esc["Piloto"]
            pts_piloto = row_esc["Puntos Finales Proyectados"]
            brecha = pts_lider - pts_piloto
            
            with st.expander(f"🏎️ Situación de {piloto_name}"):
                if piloto_name == lider_actual_escenario:
                    segundo_nombre = df_ranking_actual.loc[1, "Piloto"] if len(df_ranking_actual) > 1 else "el segundo"
                    limite_seguro = max(0.0, puntos_maximos_por_fecha - brecha_lider)
                    
                    st.write(f"👑 **Blindaje de Título:**")
                    st.write(f"- Si sumas más de **{limite_seguro:.2f} puntos**, **{segundo_nombre}** ya no tiene forma matemática de alcanzarte por más que gane todo.")
                
                else:
                    puntos_necesarios_perseguidor = brecha + puntos_ejemplo_lider
                    
                    st.write(f"🔥 **Condición de campeonato para {piloto_name}:**")
                    if brecha > puntos_maximos_por_fecha or puntos_necesarios_perseguidor > puntos_maximos_por_fecha:
                        st.write(f"- Matemáticamente **sin chances**: aun si haces la fecha perfecta, dependes de que {lider_actual_escenario} sume menos del mínimo posible.")
                    else:
                        st.write(f"- **La cuenta exacta:** Si **{lider_actual_escenario}** hace una fecha ajustada sumando por ejemplo **{puntos_ejemplo_lider:.2f} puntos**, tú estás obligado a sumar al menos **{puntos_necesarios_perseguidor:.2f} puntos** netos en el fin de semana para superarlo en la tabla final y arrebatarle el campeonato.")
        # 🏎️ PANEL VISUAL DE PROYECCIÓN GENERAL (ESTILO F1 TV - SIN TABLAS)
        # =========================================================================
        st.markdown("### 📊 Posiciones Finales Proyectadas")
        
        # --- ESTILOS CSS REFINADOS PARA LAS TARJETAS HORIZONTALES ---
        estilos_panel_proy = """
        <style>
        .contenedor-proy {
            background-color: #161925;
            border-left: 5px solid #e10600;
            border-radius: 10px;
            padding: 14px 20px;
            margin-bottom: 12px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.5);
        }
        .izquierda-proy {
            display: flex;
            align-items: center;
            gap: 20px;
        }
        .puesto-proy {
            color: #e10600;
            font-size: 20px;
            font-weight: bold;
            font-family: 'monospace';
            midth: 35px;
        }
        .nombre-proy {
            color: #ffffff;
            font-size: 16px;
            font-weight: bold;
        }
        .puntos-ganados-proy {
            color: #9fa6b2;
            font-size: 12px;
            display: block;
            margin-top: 2px;
        }
        .derecha-proy {
            text-align: right;
            display: flex;
            gap: 30px;
            align-items: center;
        }
        .bloque-metrica-proy {
            text-align: center;
            min-width: 90px;
        }
        .etiqueta-metrica-proy {
            color: #9fa6b2;
            font-size: 10px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            display: block;
            margin-bottom: 2px;
        }
        .valor-pts-proy {
            color: #ffffff;
            font-size: 18px;
            font-weight: bold;
            font-family: 'monospace';
        }
        .valor-kg-proy {
            color: #ff9100;
            font-size: 18px;
            font-weight: bold;
            font-family: 'monospace';
        }
        </style>
        """
        st.markdown(estilos_panel_proy, unsafe_allow_html=True)

        # Generamos el HTML dinámico piloto por piloto ordenados del 1º al 4º
        html_panel = ""
        for idx_p, row_p in df_proyeccion_final.iterrows():
            puesto = idx_p + 1
            nombre = row_p["Piloto"]
            ganados = row_p["Pts Sumados Simulación"]
            finales = row_p["Puntos Finales Proyectados"]
            lastre_futuro = row_p["Lastre Próxima Fecha"]
            
            # Cambia el color del texto del Kilaje dinámicamente según la severidad de la carga
            color_kg_dinamico = "#e10600" if lastre_futuro > 30.0 else ("#ff9100" if lastre_futuro > 10.0 else "#00e676")
            # El líder de la simulación se destaca con un borde verde neón en su tarjeta
            color_borde_tarjeta = "#00e676" if idx_p == 0 else "#e10600"
            
            html_panel += f"""
            <div class="contenedor-proy" style="border-left: 5px solid {color_borde_tarjeta};">
                <div class="izquierda-proy">
                    <span class="puesto-proy">#{puesto}</span>
                    <div>
                        <span class="nombre-proy">🏎️ {nombre}</span>
                        <span class="puntos-ganados-proy">+{ganados:.2f} pts en simulación</span>
                    </div>
                </div>
                <div class="derecha-proy">
                    <div class="bloque-metrica-proy">
                        <span class="etiqueta-metrica-proy">Puntaje Final</span>
                        <span class="valor-pts-proy">{finales:.2f} pts</span>
                    </div>
                    <div class="bloque-metrica-proy">
                        <span class="etiqueta-metrica-proy">Próximo Lastre</span>
                        <span class="valor-kg-proy" style="color: {color_kg_dinamico};">{lastre_futuro:.1f} Kg</span>
                    </div>
                </div>
            </div>
            """
            
        st.markdown(html_panel, unsafe_allow_html=True)

elif opcion == "Lastre":
    st.title("⚖️ Sistema de Lastre Técnico")
    st.write("Consulta el peso acumulado en pista y el rendimiento de los pilotos con lastre.")

    if os.path.exists(ARCHIVO_EXCEL):
        try:
            # 1. LEER LAS PLANILLAS DE FORMA SEGURA
            df_hoja1_sim = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')
            columna_f_sim = df_hoja1_sim.iloc[:, 5].astype(str).str.strip().str.upper()

            # Mapeo de columnas oficiales de pilotos en tu Hoja1
            indices_pilotos_hoja1 = {"Agus": 6, "Pablo": 8, "Juandi": 10, "Eze": 12}
            pilotos_torneo = list(indices_pilotos_hoja1.keys())

            filas_lastre_fecha_real = columna_f_sim[columna_f_sim.str.contains("LASTRE FECHA|LASTRE.*FECHA", na=False)].index.tolist()
            filas_lastre_acum_real = columna_f_sim[columna_f_sim.str.contains("LASTRE ACUMULADO|LASTRE.*ACUM", na=False)].index.tolist()

            if len(filas_lastre_fecha_real) > 0:
                total_fechas_detectadas = len(filas_lastre_fecha_real)
                opciones_lastre_dinamicas = [f"Fecha {i+1}" for i in range(total_fechas_detectadas)]
                fecha_sel = st.selectbox("Seleccionar Fecha para Consultar Lastre Técnico:", opciones_lastre_dinamicas)
                
                import re
                numeros_encontrados = re.findall(r'\d+', str(fecha_sel))
                idx_fecha_sel = int(numeros_encontrados[0]) - 1 if numeros_encontrados else 0

                # Armamos la tabla para renderizar los kilos en pista actuales
                tabla_lastre = []
                fila_pista = filas_lastre_acum_real[idx_fecha_sel] if idx_fecha_sel < len(filas_lastre_acum_real) else None
                fila_generado = filas_lastre_fecha_real[idx_fecha_sel] if idx_fecha_sel < len(filas_lastre_fecha_real) else None
                
                for p in pilotos_torneo:
                    col_idx = indices_pilotos_hoja1[p]
                    k_pista = df_hoja1_sim.iloc[fila_pista, col_idx] if fila_pista is not None else 0
                    k_gen = df_hoja1_sim.iloc[fila_generado, col_idx] if fila_generado is not None else 0
                    tabla_lastre.append({
                        "Piloto": p,
                        "Lastre en Pista (kg)": float(str(k_pista).upper().replace("KG","").strip()) if pd.notna(k_pista) else 0.0,
                        "Lastre Generado (kg)": float(str(k_gen).upper().replace("KG","").strip()) if pd.notna(k_gen) else 0.0
                    })
                
                df_lastre_render = pd.DataFrame(tabla_lastre)
                
                # Dibujamos las tarjetas individuales arriba
                cols_cards = st.columns(len(pilotos_torneo))
                for i, row in enumerate(df_lastre_render.iterrows()):
                    with cols_cards[i % 4]:
                        st.metric(label=f"Auto de {row[1]['Piloto']}", value=f"{row[1]['Lastre en Pista (kg)']} Kg", delta=f"+{row[1]['Lastre Generado (kg)']} Kg ganados")
            else:
                st.warning("No se detectaron las filas de 'Lastre Fecha' en la 'Hoja1'.")

        except Exception as e:
            st.error(f"Error al calcular el lastre desde el Excel: {e}")
    else:
        st.warning(f"No se encontró el archivo '{ARCHIVO_EXCEL}'.")       

elif opcion == "Estadisticas":
    st.title("🏁 Live Timing Pro — Análisis de Ritmo")

    # =========================================================================
    # FUNCIONES AUXILIARES DE FORMATO Y PROCESAMIENTO DE TIEMPOS
    # =========================================================================
    def formatear_mm_ss_ms(segundos):
        """Formatea segundos flotantes a string MM:SS,mmm"""
        if segundos is None or pd.isna(segundos):
            return "--:--"
        minutos = int(segundos // 60)
        seg_restantes = segundos % 60
        seg = int(seg_restantes)
        miliseg = int(round((seg_restantes - seg) * 1000))
        return f"{minutos:02d}:{seg:02d},{miliseg:03d}"

    def limpiar_tiempo_str(tiempo_raw):
        """Limpia y valida cadenas de tiempo raw de Excel"""
        if not tiempo_raw or tiempo_raw == "--:--":
            return "--:--"
        seg = tiempo_a_segundos(tiempo_raw)
        return formatear_mm_ss_ms(seg)

    def procesar_hoja_dinamica_multicolumnas(file_path, sheet_name):
        """Extrae vueltas y promedios desde la hoja 'Diferencia en Carrera'"""
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        registros = []
        promedios_excel = {}

        for r in range(len(df_raw)):
            for c in range(len(df_raw.columns)):
                val = str(df_raw.iloc[r, c]).strip()
                if val and val.lower() != 'nan' and not val.isdigit() and len(val) > 2:
                    val_upper = val.upper()
                    if not any(k in val_upper for k in ['PILOTO', 'PROMEDIO', 'MEJOR TIEMPO', 'CARRERA', '+', ':']):
                        circuito_nombre = val_upper.replace(' C1', '').replace(' C2', '').strip()
                        carrera_nombre = 'Carrera 1' if 'C1' in val_upper else ('Carrera 2' if 'C2' in val_upper else 'General')
                        
                        for r_sub in range(r + 1, min(r + 5, len(df_raw))):
                            for c_sub in range(max(0, c - 2), min(len(df_raw.columns) - 1, c + 8)):
                                header_piloto = str(df_raw.iloc[r_sub, c_sub]).strip()
                                if header_piloto and header_piloto.lower() != 'nan' and header_piloto.upper() not in ['PILOTO', 'PROMEDIO', 'MEJOR TIEMPO']:
                                    if not header_piloto.isdigit() and ':' not in header_piloto:
                                        # Extracción de tiempos vuelta a vuelta
                                        for r_data in range(r_sub + 1, len(df_raw)):
                                            v_val = df_raw.iloc[r_data, c_sub]
                                            t_val = df_raw.iloc[r_data, c_sub + 1]
                                            if pd.notna(v_val) and pd.notna(t_val):
                                                v_str = str(v_val).strip()
                                                if v_str.isdigit():
                                                    registros.append({
                                                        'Circuito': circuito_nombre,
                                                        'Carrera': carrera_nombre,
                                                        'Piloto': header_piloto,
                                                        'Vuelta': int(v_str),
                                                        'Tiempo': str(t_val).strip()
                                                    })
                                                else:
                                                    break
                                            else:
                                                break

                                        # Extracción de promedios fijados en el Excel
                                        for r_prom in range(r_sub + 1, len(df_raw)):
                                            fila_valores = [str(df_raw.iloc[r_prom, col]).strip().upper() for col in range(max(0, c_sub - 3), c_sub + 1)]
                                            if 'PROMEDIO' in fila_valores:
                                                t_prom = df_raw.iloc[r_prom, c_sub + 1]
                                                if pd.notna(t_prom):
                                                    promedios_excel[(circuito_nombre, carrera_nombre, header_piloto)] = str(t_prom).strip()
                                                break
        return pd.DataFrame(registros).drop_duplicates(), promedios_excel

    def filtrar_grafico_inteligente(df_circuito):
        """Filtra anomalías generales y vueltas con Pace Car (>=50% de coches lentos)"""
        if df_circuito.empty:
            return df_circuito

        ritmos_base = df_circuito.groupby('Piloto')['Tiempo_Seg'].apply(
            lambda x: x.nsmallest(5).median() if len(x) >= 3 else x.median()
        ).to_dict()

        df_temp = df_circuito.copy()
        df_temp['Ritmo_Base'] = df_temp['Piloto'].map(ritmos_base)
        df_temp['Es_Anomalia'] = df_temp['Tiempo_Seg'] > (df_temp['Ritmo_Base'] + 7.0)

        anomalias_por_vuelta = df_temp.groupby('Vuelta')['Es_Anomalia'].agg(
            total_pilotos='count',
            pilotos_anomalos='sum'
        ).reset_index()

        anomalias_por_vuelta['Es_Pace_Car'] = (
            (anomalias_por_vuelta['pilotos_anomalos'] / anomalias_por_vuelta['total_pilotos']) >= 0.5
        ) & (anomalias_por_vuelta['pilotos_anomalos'] > 0)

        vueltas_pace_car = set(anomalias_por_vuelta[anomalias_por_vuelta['Es_Pace_Car']]['Vuelta'])
        return df_circuito[~df_circuito['Vuelta'].isin(vueltas_pace_car)].copy()

    # =========================================================================
    # EJECUCIÓN PRINCIPAL Y CÁLCULO DE TELEMETRÍA
    # =========================================================================
    if os.path.exists(ARCHIVO_EXCEL):
        try:
            df_est, promedios_excel = procesar_hoja_dinamica_multicolumnas(ARCHIVO_EXCEL, "Diferencia en Carrera")

            if not df_est.empty:
                df_est['Tiempo_Seg'] = df_est['Tiempo'].apply(tiempo_a_segundos)
                df_validos = df_est.dropna(subset=['Tiempo_Seg']).copy()

                # 1. SELECCIÓN DE CIRCUITO Y CARRERA
                circuitos_est = sorted(list(df_validos['Circuito'].unique()))
                circ_sel = st.selectbox("🏎️ Seleccionar Circuito:", circuitos_est)

                df_circ = df_validos[df_validos['Circuito'] == circ_sel].sort_values('Vuelta')
                cargas_disponibles = sorted(list(df_circ['Carrera'].unique()))
                
                carrera_sel = "Todas las Carreras"
                if len(cargas_disponibles) > 1:
                    opciones_carrera = ["Todas las Carreras"] + cargas_disponibles
                    carrera_sel = st.radio("🏁 Seleccionar Carrera:", opciones_carrera, horizontal=True)
                    if carrera_sel != "Todas las Carreras":
                        df_circ = df_circ[df_circ['Carrera'] == carrera_sel]

                pilotos_disponibles = sorted(list(df_circ['Piloto'].unique()))
                
                # 2. COMPARATIVA FRENTE A FRENTE EN SIDEBAR
                st.sidebar.markdown("---")
                st.sidebar.subheader("Comparativa Frente a Frente")
                p1 = st.sidebar.selectbox("Piloto 1:", pilotos_disponibles, index=0 if len(pilotos_disponibles) > 0 else 0)
                p2_idx = 1 if len(pilotos_disponibles) > 1 else 0
                p2 = st.sidebar.selectbox("Piloto 2:", pilotos_disponibles, index=p2_idx)

                df_p1 = df_circ[df_circ['Piloto'] == p1]
                df_p2 = df_circ[df_circ['Piloto'] == p2]

                # Métricas de Piloto 1
                rec_p1 = df_p1['Tiempo_Seg'].min() if not df_p1.empty else None
                rec_p1_str = limpiar_tiempo_str(df_p1.loc[df_p1['Tiempo_Seg'].idxmin(), 'Tiempo']) if rec_p1 else "--:--"

                # Métricas de Piloto 2
                rec_p2 = df_p2['Tiempo_Seg'].min() if not df_p2.empty else None
                rec_p2_str = limpiar_tiempo_str(df_p2.loc[df_p2['Tiempo_Seg'].idxmin(), 'Tiempo']) if rec_p2 else "--:--"

                carrera_actual = df_circ['Carrera'].iloc[0] if carrera_sel != "Todas las Carreras" and not df_circ.empty else 'Carrera 1'
                
                prom_p1_str_raw = promedios_excel.get((circ_sel, carrera_actual, p1))
                prom_p2_str_raw = promedios_excel.get((circ_sel, carrera_actual, p2))

                prom_p1 = tiempo_a_segundos(prom_p1_str_raw) if prom_p1_str_raw else (df_p1['Tiempo_Seg'].mean() if not df_p1.empty else None)
                prom_p2 = tiempo_a_segundos(prom_p2_str_raw) if prom_p2_str_raw else (df_p2['Tiempo_Seg'].mean() if not df_p2.empty else None)

                str_prom_p1 = formatear_mm_ss_ms(prom_p1)
                str_prom_p2 = formatear_mm_ss_ms(prom_p2)

                # Desviación / Regularidad (Vuelta limpia <= 105% del récord personal)
                def filtrar_regularidad_piloto(df_piloto):
                    if df_piloto.empty:
                        return df_piloto
                    record_personal = df_piloto['Tiempo_Seg'].min()
                    return df_piloto[df_piloto['Tiempo_Seg'] <= record_personal * 1.05]

                df_p1_reg = filtrar_regularidad_piloto(df_p1)
                df_p2_reg = filtrar_regularidad_piloto(df_p2)

                std_p1 = df_p1_reg['Tiempo_Seg'].std() if len(df_p1_reg) > 1 else None
                std_p2 = df_p2_reg['Tiempo_Seg'].std() if len(df_p2_reg) > 1 else None

                if std_p1 is not None and std_p2 is not None:
                    if std_p1 < std_p2:
                        txt_reg_p1 = f"±{std_p1:.3f}s (MÁS SÓLIDO)"
                        txt_reg_p2 = f"±{std_p2:.3f}s"
                    else:
                        txt_reg_p1 = f"±{std_p1:.3f}s"
                        txt_reg_p2 = f"±{std_p2:.3f}s (MÁS SÓLIDO)"
                else:
                    txt_reg_p1 = f"±{std_p1:.3f}s" if std_p1 is not None else "--"
                    txt_reg_p2 = f"±{std_p2:.3f}s" if std_p2 is not None else "--"

                # 3. TARJETAS DE PILOTOS Y DUELOS DIRECTOS
                st.markdown("<br>", unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                
                with c1:
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #161925 0%, #0d0f17 100%); padding: 18px; border-radius: 12px; border-top: 4px solid #00d2ff; box-shadow: 0 4px 12px rgba(0,0,0,0.5); margin-bottom: 12px; text-align: center;">
                        <span style="color: #8f92a1; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;">🏎️ PILOTO 1</span>
                        <h2 style="color: #ffffff; font-size: 26px; font-weight: 800; margin: 4px 0 12px 0;">{p1}</h2>
                        <p style="color: #e0e0e0; font-size: 15px; margin: 4px 0;">⏱️ <b style="color: #ffffff;">Récord:</b> <span style="color: #00d2ff; font-weight: 700;">{rec_p1_str}</span></p>
                        <p style="color: #e0e0e0; font-size: 15px; margin: 4px 0;">📊 <b style="color: #ffffff;">Promedio (Excel):</b> <span style="color: #00d2ff; font-weight: 700;">{str_prom_p1}</span></p>
                        <p style="color: #e0e0e0; font-size: 14px; margin: 4px 0;">🎯 <b style="color: #ffffff;">Regularidad:</b> <span style="color: #ff007f; font-weight: 700;">{txt_reg_p1}</span></p>
                    </div>
                    """, unsafe_allow_html=True)

                with c2:
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #161925 0%, #0d0f17 100%); padding: 18px; border-radius: 12px; border-top: 4px solid #ff9900; box-shadow: 0 4px 12px rgba(0,0,0,0.5); margin-bottom: 12px; text-align: center;">
                        <span style="color: #8f92a1; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;">🏎️ PILOTO 2</span>
                        <h2 style="color: #ffffff; font-size: 26px; font-weight: 800; margin: 4px 0 12px 0;">{p2}</h2>
                        <p style="color: #e0e0e0; font-size: 15px; margin: 4px 0;">⏱️ <b style="color: #ffffff;">Récord:</b> <span style="color: #ff9900; font-weight: 700;">{rec_p2_str}</span></p>
                        <p style="color: #e0e0e0; font-size: 15px; margin: 4px 0;">📊 <b style="color: #ffffff;">Promedio (Excel):</b> <span style="color: #ff9900; font-weight: 700;">{str_prom_p2}</span></p>
                        <p style="color: #e0e0e0; font-size: 14px; margin: 4px 0;">🎯 <b style="color: #ffffff;">Regularidad:</b> <span style="color: #ff007f; font-weight: 700;">{txt_reg_p2}</span></p>
                    </div>
                    """, unsafe_allow_html=True)

                c3, c4 = st.columns(2)
                if rec_p1 and rec_p2:
                    diff_rec = rec_p1 - rec_p2
                    ganador_rec = p1 if diff_rec < 0 else p2
                    rival_rec = p2 if diff_rec < 0 else p1
                    val_rec_str = f"{ganador_rec.upper()} (-{abs(diff_rec):.3f}s vs {rival_rec.upper()})"
                else:
                    val_rec_str = "Datos insuficientes"

                with c3:
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #2a0845 0%, #16002c 100%); padding: 16px; border-radius: 12px; border: 1px solid #8e44ad; box-shadow: 0 4px 12px rgba(0,0,0,0.5); margin-bottom: 15px; text-align: center;">
                        <span style="color: #d1a8ff; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;">⚡ VUELTA RÁPIDA DE LA TANDA</span>
                        <h3 style="color: #ffffff; font-size: 20px; font-weight: 800; margin: 6px 0 0 0;">🥇 {val_rec_str}</h3>
                    </div>
                    """, unsafe_allow_html=True)

                if prom_p1 and prom_p2:
                    diff_prom = prom_p1 - prom_p2
                    ganador_prom = p1 if diff_prom < 0 else p2
                    rival_prom = p2 if diff_prom < 0 else p1
                    val_prom_str = f"{ganador_prom.upper()} (-{abs(diff_prom):.3f}s vs {rival_prom.upper()})"
                else:
                    val_prom_str = "Datos insuficientes"

                with c4:
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #0a2e38 0%, #04161c 100%); padding: 16px; border-radius: 12px; border: 1px solid #27ae60; box-shadow: 0 4px 12px rgba(0,0,0,0.5); margin-bottom: 15px; text-align: center;">
                        <span style="color: #80e8ab; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;">👑 LÍDER DE RITMO GLOBAL</span>
                        <h3 style="color: #ffffff; font-size: 20px; font-weight: 800; margin: 6px 0 0 0;">🥇 {val_prom_str}</h3>
                    </div>
                    """, unsafe_allow_html=True)

                # =========================================================================
                # 4. GRÁFICO DE EVOLUCIÓN CON PLOTLY (ANCHO COMPLETO)
                # =========================================================================
                st.markdown("---")
                st.subheader("📈 Gráfico de Evolución de Ritmo")

                # 1. Contenedor independiente para asegurar ancho completo
                with st.container():
                    limpiar_outliers = st.checkbox("🧹 Filtrar vueltas lentas / Pace Car", value=False)
                    st.caption("💡 **¿Para qué sirve?** Detecta y remueve vueltas con incidentes o Pace Car para hacer 'zoom' en el eje Y y analizar el ritmo limpio de carrera.")

                    # Selección de datos según el switch
                    if limpiar_outliers:
                        df_grafico = filtrar_grafico_inteligente(df_circ)
                        eje_y = 'Tiempo_Fecha'
                        etiqueta_y = 'Tiempo de Vuelta'
                    else:
                        df_grafico = df_circ.copy()
                        eje_y = 'Tiempo_Seg'
                        etiqueta_y = 'Segundos de Carrera'

                    df_grafico['Tiempo_Fecha'] = pd.to_datetime(df_grafico['Tiempo_Seg'], unit='s')
                    df_grafico['Tiempo_Formateado'] = df_grafico['Tiempo_Seg'].apply(formatear_mm_ss_ms)

                    # Gráfico de Plotly
                    fig_line = px.line(
                        df_grafico,
                        x='Vuelta',
                        y=eje_y,
                        color='Piloto',
                        markers=True,
                        labels={eje_y: etiqueta_y, 'Vuelta': 'Número de Vuelta'},
                        template="plotly_dark",
                        hover_data={'Tiempo_Fecha': False, 'Tiempo_Seg': False, 'Tiempo_Formateado': True}
                    )

                    formato_y = '%M:%S,%3f' if limpiar_outliers else None

                    fig_line.update_layout(
                        height=520,
                        autosize=True,
                        margin=dict(l=20, r=20, t=30, b=20),
                        hovermode="x unified",
                        xaxis=dict(showgrid=True, gridcolor="#2a2e3d", dtick=1),
                        yaxis=dict(showgrid=True, gridcolor="#2a2e3d", tickformat=formato_y),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )

                    # Renderizado forzado a ancho completo
                    st.plotly_chart(fig_line, use_container_width=True)

                # 5. TABLA COMPARATIVA
                st.subheader("📋 Historial de Tiempos Vuelta a Vuelta")
                df_tabla_base = df_circ[['Vuelta', 'Piloto', 'Tiempo_Seg']].drop_duplicates(subset=['Vuelta', 'Piloto']).copy()
                df_tabla_base['Tiempo'] = df_tabla_base['Tiempo_Seg'].apply(formatear_mm_ss_ms)
                
                df_tabla = df_tabla_base.pivot(index='Vuelta', columns='Piloto', values='Tiempo').reset_index()
                df_tabla.columns.name = None
                df_tabla.insert(0, 'Posición', "Vuelta " + df_tabla['Vuelta'].astype(str))
                
                st.dataframe(df_tabla.drop(columns=['Vuelta']), use_container_width=True, hide_index=True)

                # 6. ANÁLISIS DE RENDIMIENTO CON LASTRE TÉCNICO
                st.markdown("---")
                st.subheader("🏋️ Análisis de Rendimiento con Lastre Técnico")
                st.write("Evaluación histórica del promedio de puntos sumados en relación a los Kg cargados:")

                df_hoja1_sim = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')
                columna_f_sim = df_hoja1_sim.iloc[:, 5].astype(str).str.strip().str.upper()
                
                df_puntos_aux = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Tabla Final', engine='openpyxl')
                df_puntos_aux.columns = [str(c).strip() for c in df_puntos_aux.columns]

                indices_pilotos_hoja1 = {"Agus": 6, "Pablo": 8, "Juandi": 10, "Eze": 12}
                pilotos_torneo = list(indices_pilotos_hoja1.keys())

                columnas_fechas_reales = [col for col in df_puntos_aux.columns if str(col).strip().upper().startswith("FECHA")]
                ultima_fecha_real_num = 0
                for idx_f, col_f in enumerate(columnas_fechas_reales):
                    if df_puntos_aux[col_f].astype(float).sum() > 0:
                        ultima_fecha_real_num = idx_f + 1
                if ultima_fecha_real_num == 0:
                    ultima_fecha_real_num = 7

                filas_lastre_fecha_real = columna_f_sim[columna_f_sim.str.contains("LASTRE FECHA|LASTRE.*FECHA", na=False)].index.tolist()
                filas_lastre_acum_real = columna_f_sim[columna_f_sim.str.contains("LASTRE ACUMULADO|LASTRE.*ACUM", na=False)].index.tolist()

                rangos_peso = ["0-20 Kg", "21-40 Kg", "41-60 Kg", "61-80 Kg+"]
                historial_peso_puntos = {p: {r: [] for r in rangos_peso} for p in pilotos_torneo}

                for idx_c, fila_total in enumerate(filas_lastre_fecha_real):
                    for piloto, col_idx in indices_pilotos_hoja1.items():
                        val_puntos = df_hoja1_sim.iloc[fila_total, col_idx]
                        puntos_limpios = str(val_puntos).replace(',', '.', 1) if pd.notna(val_puntos) else "0.0"
                        
                        # Validar si hay puntos registrados en esta fecha
                        try:
                            pts_fecha = float(puntos_limpios)
                        except ValueError:
                            pts_fecha = 0.0

                        # Si la fecha no se ha corrido/cargado aún (0 puntos o celda vacía), se omite
                        if pd.isna(val_puntos) or str(val_puntos).strip() in ['', 'nan', 'None']:
                            continue

                        kilos_largada = 0.0
                        if idx_c > 0:
                            idx_bloque_anterior = idx_c - 1
                            if idx_bloque_anterior < len(filas_lastre_acum_real):
                                fila_lastre_anterior = filas_lastre_acum_real[idx_bloque_anterior]
                                val_lastre = df_hoja1_sim.iloc[fila_lastre_anterior, col_idx]
                                if pd.notna(val_lastre):
                                    k_str = str(val_lastre).upper().replace("KG", "").replace(',', '.').strip()
                                    try:
                                        kilos_largada = float(k_str)
                                    except ValueError:
                                        kilos_largada = 0.0

                        if kilos_largada <= 20.0: rango_elegido = "0-20 Kg"
                        elif kilos_largada <= 40.0: rango_elegido = "21-40 Kg"
                        elif kilos_largada <= 60.0: rango_elegido = "41-60 Kg"
                        else: rango_elegido = "61-80 Kg+"
                            
                        historial_peso_puntos[piloto][rango_elegido].append(pts_fecha)

                estilos_mapa_peso = """
                <style>
                .contenedor-peso { background-color: #1a1c23; border-top: 4px solid #ff1801; border-radius: 8px; padding: 15px; margin-bottom: 20px; box-shadow: 2px 2px 10px rgba(0,0,0,0.4); }
                .titulo-peso { color: #ff1801; font-size: 16px; font-weight: bold; text-align: center; margin-bottom: 12px; text-transform: uppercase; letter-spacing: 1px; }
                .fila-piloto-peso { display: flex; justify-content: space-between; align-items: center; padding: 6px 0; border-bottom: 1px solid #2d313f; }
                .fila-piloto-peso:last-child { border-bottom: none; }
                .nombre-piloto-peso { color: #ffffff; font-weight: 500; font-size: 14px; }
                .datos-piloto-peso { text-align: right; }
                .pts-peso { color: #ffffff; font-weight: bold; font-size: 15px; font-family: 'monospace'; }
                .carreras-peso { color: #8a8d9a; font-size: 11px; display: block; }
                </style>
                """
                st.markdown(estilos_mapa_peso, unsafe_allow_html=True)
                columnas_rangos = st.columns(4)

                for idx_rango, rango in enumerate(rangos_peso):
                    with columnas_rangos[idx_rango]:
                        html_bloque = '<div class="contenedor-peso">'
                        html_bloque += '<div class="titulo-peso">📦 ' + str(rango) + '</div>'
                        
                        tabla_pilotos_rango = []
                        for piloto in pilotos_torneo:
                            lista_puntos = historial_peso_puntos[piloto][rango]
                            cant_fechas = len(lista_puntos)
                            if cant_fechas > 0:
                                promedio = sum(lista_puntos) / cant_fechas
                                txt_fechas = f"en {cant_fechas} GP" if cant_fechas > 1 else "en 1 GP"
                                tabla_pilotos_rango.append({"nombre": piloto, "promedio": promedio, "txt_pts": f"{promedio:.1f} pts", "txt_fechas": txt_fechas})
                            else:
                                tabla_pilotos_rango.append({"nombre": piloto, "promedio": -1.0, "txt_pts": "-", "txt_fechas": "0 GP"})
                        
                        tabla_pilotos_rango = sorted(tabla_pilotos_rango, key=lambda x: x["promedio"], reverse=True)
                        for p_data in tabla_pilotos_rango:
                            html_bloque += '<div class="fila-piloto-peso">'
                            html_bloque += '    <span class="nombre-piloto-peso">🏎️ ' + str(p_data["nombre"]) + '</span>'
                            html_bloque += '    <div class="datos-piloto-peso">'
                            html_bloque += '        <span class="pts-peso">' + str(p_data["txt_pts"]) + '</span>'
                            html_bloque += '        <span class="carreras-peso">' + str(p_data["txt_fechas"]) + '</span>'
                            html_bloque += '    </div>'
                            html_bloque += '</div>'
                        
                        html_bloque += '</div>'
                        st.markdown(html_bloque, unsafe_allow_html=True)
                st.info("💡 **Análisis de Telemetría:** Las tarjetas ordenan automáticamente a los pilotos de mayor a menor efectividad dentro de cada rango de Kg.")

                # =========================================================================
                # 🕸️ RADAR MULTIVARIABLE DE ESTADÍSTICAS (CÁLCULO DE FECHAS CORREGIDO)
                # =========================================================================
                st.markdown("---")
                st.subheader("🕸️ Perfil Comparativo Multivariable")
                st.caption("Haz clic en los nombres de los pilotos en la leyenda para activar o desactivar su perfil.")

                try:
                    # 1. Leer el Excel directamente
                    df_raw = pd.read_excel(ARCHIVO_EXCEL, header=None)
                    #st.write(f"Archivo cargado. Dimensiones: {df_raw.shape}")
                    # Mapeo de columnas según matriz de Excel
                    mapa_pilotos = {
                        "Agus": {"col_val": 6, "col_pts": 7},
                        "Pablo": {"col_val": 8, "col_pts": 9},
                        "Juandi": {"col_val": 10, "col_pts": 11},
                        "Eze": {"col_val": 12, "col_pts": 13}
                    }

                    # Identificar filas clave en Columna F (idx 5)
                    col_f = df_raw.iloc[:, 5].astype(str).str.strip().str.upper()

                    filas_c1 = df_raw[col_f == "C1"]
                    filas_c2 = df_raw[col_f == "C2"]
                    filas_pole = df_raw[col_f == "POLE"]

                    # =========================================================================
                    # BUCLE DE CÁLCULO DE ESTADÍSTICAS (ROBUSTO)
                    # =========================================================================
                    datos_radar = []

                    for piloto in pilotos_torneo:
                        if piloto not in mapa_pilotos:
                            continue

                        idx_v = mapa_pilotos[piloto]["col_val"]
                        idx_p = mapa_pilotos[piloto]["col_pts"]

                        # 1. Extraer posiciones válidas de C1 y C2
                        vals_c1_raw = pd.to_numeric(filas_c1.iloc[:, idx_v], errors='coerce')
                        vals_c2_raw = pd.to_numeric(filas_c2.iloc[:, idx_v], errors='coerce')

                        pos_c1 = [int(x) for x in vals_c1_raw if pd.notna(x) and int(x) in [1, 2, 3, 4]]
                        pos_c2 = [int(x) for x in vals_c2_raw if pd.notna(x) and int(x) in [1, 2, 3, 4]]
                        
                        # 2. DEFINICIÓN DE DENOMINADORES
                        # Usamos C1 como base para las fechas del torneo (asumiendo que todos corren C1)
                        fechas_disputadas = len(pos_c1) if len(pos_c1) > 0 else 1
                        
                        todas_pos = pos_c1 + pos_c2
                        total_mangas = len(todas_pos) if len(todas_pos) > 0 else 1
                        
                        # Cálculo de posición promedio
                        pos_prom = (sum(todas_pos) / total_mangas) if total_mangas > 0 else 2.5
                        
                        # Poles totales
                        poles_fechas = pd.to_numeric(filas_pole.iloc[:, idx_p], errors='coerce').fillna(0)
                        poles_totales = int((poles_fechas == 1).sum())

                        # 3. CÁLCULO DE MÉTRICAS
                        pct_conversion_pole = (poles_totales / fechas_disputadas) * 100.0

                        # Pole a P1 en C1
                        pole_y_p1_count = 0
                        for idx_f in filas_pole.index:
                            val_pole = pd.to_numeric(df_raw.iloc[idx_f, idx_p], errors='coerce')
                            idx_c1_asoc = idx_f - 2
                            val_c1 = pd.to_numeric(df_raw.iloc[idx_c1_asoc, idx_v], errors='coerce') if idx_c1_asoc in df_raw.index else None
                            if val_pole == 1 and val_c1 == 1:
                                pole_y_p1_count += 1

                        pct_pole_a_p1 = (pole_y_p1_count / poles_totales * 100.0) if poles_totales > 0 else 0.0
                        pct_ritmo = max(0.0, ((5.0 - pos_prom) / 4.0) * 100.0)
                        
                        cant_podios = sum(1 for p in todas_pos if p in [1, 2, 3])
                        pct_podios = (cant_podios / total_mangas) * 100.0
                        
                        cant_victorias = sum(1 for p in todas_pos if p == 1)
                        pct_victorias = (cant_victorias / total_mangas) * 100.0

                        datos_radar.append({
                            "Piloto": piloto,
                            "pos_prom": pos_prom,
                            "poles_real": poles_totales,
                            "pct_ritmo": pct_ritmo,
                            "pct_podios": pct_podios,
                            "cant_podios": cant_podios,
                            "cant_victorias": cant_victorias,
                            "pct_victorias": pct_victorias,
                            "total_mangas": total_mangas,
                            "fechas_disputadas": fechas_disputadas,
                            "pct_conversion_pole": pct_conversion_pole,
                            "pct_pole_a_p1": pct_pole_a_p1
                        })

                    for piloto in pilotos_torneo:
                        if piloto not in mapa_pilotos:
                            continue

                        idx_v = mapa_pilotos[piloto]["col_val"]
                        idx_p = mapa_pilotos[piloto]["col_pts"]

                        # Extraer posiciones válidas
                        vals_c1_raw = pd.to_numeric(filas_c1.iloc[:, idx_v], errors='coerce')
                        vals_c2_raw = pd.to_numeric(filas_c2.iloc[:, idx_v], errors='coerce')

                        pos_c1 = [int(x) for x in vals_c1_raw if pd.notna(x) and int(x) in [1, 2, 3, 4]]
                        pos_c2 = [int(x) for x in vals_c2_raw if pd.notna(x) and int(x) in [1, 2, 3, 4]]
                        
                        # Lógica corregida para denominadores iguales
                        fechas_disputadas = len(pos_c1) if len(pos_c1) > 0 else 1
                        todas_pos = pos_c1 + pos_c2
                        total_mangas = len(todas_pos) if len(todas_pos) > 0 else 1
                        
                        # Cálculo de posición promedio (necesaria para el % Ritmo)
                        pos_prom = (sum(todas_pos) / total_mangas) if len(todas_pos) > 0 else 2.5
                        
                        # Poles totales del piloto
                        poles_fechas = pd.to_numeric(filas_pole.iloc[:, idx_p], errors='coerce').fillna(0)
                        poles_totales = int((poles_fechas == 1).sum())

                        # 1. % Conversión Pole
                        pct_conversion_pole = (poles_totales / fechas_disputadas) * 100.0

                        # 2. % Pole a P1 en C1
                        pole_y_p1_count = 0
                        for idx_f in filas_pole.index:
                            val_pole = pd.to_numeric(df_raw.iloc[idx_f, idx_p], errors='coerce')
                            idx_c1_asoc = idx_f - 2
                            val_c1 = pd.to_numeric(df_raw.iloc[idx_c1_asoc, idx_v], errors='coerce') if idx_c1_asoc in df_raw.index else None
                            if val_pole == 1 and val_c1 == 1:
                                pole_y_p1_count += 1

                        pct_pole_a_p1 = (pole_y_p1_count / poles_totales * 100.0) if poles_totales > 0 else 0.0

                        # 3. % Ritmo Carrera
                        pct_ritmo = max(0.0, ((5.0 - pos_prom) / 4.0) * 100.0)

                        # 4. % Podios
                        cant_podios = sum(1 for p in todas_pos if p in [1, 2, 3])
                        pct_podios = (cant_podios / total_mangas) * 100.0

                        # 5. % Victorias
                        cant_victorias = sum(1 for p in todas_pos if p == 1)
                        pct_victorias = (cant_victorias / total_mangas) * 100.0
                        datos_radar.append({
                            "Piloto": piloto,
                            "pos_prom": pos_prom,
                            "poles_real": poles_totales,
                            "pct_ritmo": pct_ritmo,
                            "pct_podios": pct_podios,
                            "cant_podios": cant_podios,
                            "cant_victorias": cant_victorias,
                            "pct_victorias": pct_victorias,
                            "total_mangas": total_mangas,
                            "fechas_disputadas": fechas_disputadas,
                            "pct_conversion_pole": pct_conversion_pole,
                            "pct_pole_a_p1": pct_pole_a_p1
                        })
                    filas_plotly = []
                    for d in datos_radar:
                        piloto = d["Piloto"]

                        metricas = [
                            ("% Pole a P1 (C1)", d["pct_pole_a_p1"], f"{d['pct_pole_a_p1']:.1f}%"),
                            ("% Podios", d["pct_podios"], f"{d['pct_podios']:.1f}% ({d['cant_podios']}/{d['total_mangas']} mangas)"),
                            ("% Victorias", d["pct_victorias"], f"{d['pct_victorias']:.1f}% ({d['cant_victorias']}/{d['total_mangas']} victorias)"),
                            ("% Conversión Pole", d["pct_conversion_pole"], f"{d['pct_conversion_pole']:.1f}% ({d['poles_real']}/{d['fechas_disputadas']} fechas)"),
                            ("% Ritmo Carrera", d["pct_ritmo"], f"P{d['pos_prom']:.1f} Prom.")
                        ]

                        for eje, val_real_pct, txt_hover in metricas:
                            filas_plotly.append({
                                "Piloto": piloto,
                                "Métrica": eje,
                                "Valor_Norm": max(0.0, float(val_real_pct)),
                                "Valor_Real": txt_hover
                            })

                    df_radar = pd.DataFrame(filas_plotly)

                    # Cierre de polígonos
                    listas_cerradas = []
                    for piloto in pilotos_torneo:
                        df_p_radar = df_radar[df_radar["Piloto"] == piloto].copy()
                        if not df_p_radar.empty:
                            primera_fila = df_p_radar.iloc[[0]].copy()
                            df_p_radar = pd.concat([df_p_radar, primera_fila], ignore_index=True)
                            listas_cerradas.append(df_p_radar)

                    if listas_cerradas:
                        df_radar = pd.concat(listas_cerradas, ignore_index=True)

                    col_izq_r, col_centro_r, col_der_r = st.columns([0.2, 3, 0.2])

                    with col_centro_r:
                        fig_radar = px.line_polar(
                            df_radar,
                            r="Valor_Norm",
                            theta="Métrica",
                            color="Piloto",
                            line_close=True,
                            template="plotly_dark",
                            color_discrete_map={
                                "Agus": "#e10600",
                                "Pablo": "#00b0ff",
                                "Eze": "#ff9100",
                                "Juandi": "#00e676"
                            },
                            custom_data=["Valor_Real", "Piloto"]
                        )

                        # Configuración visual e interacción limpia con el mouse
                        fig_radar.update_traces(
                            fill='toself',
                            opacity=0.35,
                            line=dict(width=3),
                            marker=dict(size=8),
                            hoveron='points+fills', 
                            hovertemplate="🏎️ <b>%{customdata[1]}</b><br>📌 <b>%{theta}:</b> %{customdata[0]}<extra></extra>"
                        )

                        fig_radar.update_layout(
                            paper_bgcolor="#0f111a",
                            plot_bgcolor="#161925",
                            hovermode="closest",
                            polar=dict(
                                bgcolor="#161925",
                                radialaxis=dict(
                                    visible=True,
                                    range=[0, 105],
                                    showticklabels=False,
                                    linecolor="#2a2f45",
                                    gridcolor="#2a2f45"
                                ),
                                angularaxis=dict(
                                    gridcolor="#2a2f45",
                                    linecolor="#2a2f45",
                                    tickfont=dict(size=12, color="#ffffff")
                                )
                            ),
                            margin=dict(l=40, r=40, t=20, b=30),
                            height=450,
                            showlegend=True,
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=-0.22,
                                xanchor="center",
                                x=0.5,
                                font=dict(size=12, color="#ffffff")
                            )
                        )

                        # Renderizado del gráfico
                        st.plotly_chart(fig_radar, use_container_width=True, key="radar_final_correcto")

                        # Explicación desplegable debajo del gráfico
                        with st.expander("ℹ️ ¿Qué significa cada métrica del gráfico radar?"):
                            st.markdown("""
                            * **% Pole a P1 (C1):** Efectividad para convertir una Pole Position en victoria directa en la Carrera 1.
                            * **% Podios:** Porcentaje de mangas disputadas (C1 y C2) en las que el piloto finalizó dentro del Top 3 (P1, P2 o P3).
                            * **% Victorias:** Porcentaje de mangas ganadas (P1) sobre el total de competencias disputadas.
                            * **% Conversión Pole:** Porcentaje de fechas disputadas en las que el piloto logró quedarse con la Pole Position.
                            * **% Ritmo Carrera:** Desempeño según posición promedio final en todas las mangas (P1 = 100%, P4 = 25%).
                            """)

                except Exception as e_radar:
                    st.warning(f"No se pudo cargar el radar multivariable: {e_radar}")

            else:
                st.warning("No se detectaron datos en la pestaña 'Diferencia en Carrera'.")

        except Exception as e:
            st.error(f"Error al generar las estadísticas: {e}.")
    else:
        st.error(f"No se encontró el archivo '{ARCHIVO_EXCEL}'.")