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
    ["Resumen", "Comparativa de Tiempos", "Lastre", "Duelo H2H", "Simulador de Campeonato"]
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
            # 1. CARGA DE BASE DE DATOS Y FILTRADOS DE CONTROL
            df_puntos_graf = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Tabla Final', engine='openpyxl')
            df_puntos_graf.columns = [str(c).strip() for c in df_puntos_graf.columns]
            df_puntos_graf = df_puntos_graf.dropna(subset=["PILOTO"])

            df_hoja1_graf = pd.read_excel(ARCHIVO_EXCEL, sheet_name='Hoja1', header=None, engine='openpyxl')

            
            # Limpiamos la columna F de tu Hoja1 (donde dice C1, C2, Pole, etc.)
            columna_f_graf = df_hoja1_graf.iloc[:, 5].astype(str).str.strip().str.upper().str.replace("Ó", "O", regex=False)
            
            filas_totales_fecha = columna_f_graf[columna_f_graf == "TOTAL FECHA"].index.tolist()
            filas_lastre_acum = columna_f_graf[columna_f_graf == "LASTRE ACUMULADO"].index.tolist()
            # Buscamos C1, C2 y Pole de forma exacta sin importar espacios fantasmas en las celdas
            filas_c1 = columna_f_graf[columna_f_graf == "C1"].index.tolist()
            filas_c2 = columna_f_graf[columna_f_graf == "C2"].index.tolist()
            
            # 🛠️ FILTRO TOLERANTE: Captura "Pole", "Pole " o cualquier variante en tu Columna F
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
                    if pd.notna(val_puntos) and float(str(val_puntos).replace(',','.',1)) > 0:
                        tiene_datos_esta_fecha = True
                if tiene_datos_esta_fecha:
                    ultima_fecha_con_datos = idx + 1

            fecha_seleccionada = st.selectbox("Seleccionar Fecha o Histórico:", opciones_fechas_combinadas)

            if fecha_seleccionada == "Campeonato Completo":
                df_filtrado_resumen = df_puntos_graf[["PILOTO", "PTS"]].sort_values(by="PTS", ascending=False).reset_index(drop=True)
                df_filtrado_resumen.columns = ["Piloto", "Puntos"]
                titulo_grafico = "Puntos - Campeonato Completo"
            else:
                df_filtrado_resumen = df_puntos_graf[["PILOTO", fecha_seleccionada]].sort_values(by=fecha_seleccionada, ascending=False).reset_index(drop=True)
                df_filtrado_resumen.columns = ["Piloto", "Puntos"]
                titulo_grafico = f"Puntos - {fecha_seleccionada}"

            # --- TABLA Y GRÁFICO SUPERIOR ---
            col1, col2 = st.columns(2)
            with col1:
                st.subheader(f"📋 Tabla de Posiciones ({fecha_seleccionada})")
                st.dataframe(
                    df_filtrado_resumen,
                    column_config={
                        "Piloto": st.column_config.TextColumn("Piloto 🏎️", width="medium"),
                        "Puntos": st.column_config.ProgressColumn(
                            "Puntos Generales",
                            format="%.2f pts",
                            min_value=0,
                            max_value=int(df_filtrado_resumen["Puntos"].max() if len(df_filtrado_resumen) > 0 else 300)
                        )
                    },
                    use_container_width=True,
                    hide_index=True
                )

            with col2:
                st.subheader("📊 Gráfico de Rendimiento")
                fig_barras = px.bar(
                    df_filtrado_resumen, x="Piloto", y="Puntos", color="Piloto",
                    template="plotly_dark", title=titulo_grafico,
                    color_discrete_sequence=["#e10600", "#00b0ff", "#ff9100", "#00e676"]
                )
                fig_barras.update_layout(plot_bgcolor="#161925", paper_bgcolor="#0f111a", margin=dict(l=10, r=10, t=40, b=10), showlegend=False)
                st.plotly_chart(fig_barras, use_container_width=True, key="resumen_rendimiento_fijo")

            # --- GRÁFICO DE LÍNEAS (EVOLUCIÓN) ---
            for idx, fila_total in enumerate(filas_totales_fecha[:ultima_fecha_con_datos]):
                nombre_fecha_eje_x = f"Fecha {idx + 1}"
                try:
                    circuito_real = opciones_fechas_combinadas[idx + 1].split(" - ")[1]
                except:
                    circuito_real = f"Carrera {idx + 1}"
                
                for piloto, col_idx in indices_pilotos_graf.items():
                    val_puntos_fecha = df_hoja1_graf.iloc[fila_total, col_idx]
                    puntos_fecha_limpio = str(val_puntos_fecha).replace(',', '.', 1) if pd.notna(val_puntos_fecha) else "0.0"
                    puntos_fecha = float(puntos_fecha_limpio) if puntos_fecha_limpio.replace('.','',1).isdigit() else 0.0
                    
                    puntos_acumulados_carrera[piloto] += puntos_fecha
                    if puntos_acumulados_carrera[piloto] > max_puntaje_detectado:
                        max_puntaje_detectado = puntos_acumulados_carrera[piloto]
                    
                    pos_c1 = "-"
                    pos_c2 = "-"
                    if idx < len(filas_c1):
                        val_c1 = df_hoja1_graf.iloc[filas_c1[idx], col_idx]
                        if pd.notna(val_c1): pos_c1 = f"P{int(float(str(val_c1).replace(',','.',1)))}"
                    if idx < len(filas_c2):
                        val_c2 = df_hoja1_graf.iloc[filas_c2[idx], col_idx]
                        if pd.notna(val_c2): pos_c2 = f"P{int(float(str(val_c2).replace(',','.',1)))}"
                    
                    resultado_txt = f"{pos_c1} / {pos_c2}"

                    lastre_txt = "0 Kg"
                    if idx > 0:
                        idx_bloque_anterior = idx - 1
                        if idx_bloque_anterior < len(filas_lastre_acum):
                            fila_lastre_anterior = filas_lastre_acum[idx_bloque_anterior]
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
                    hovertemplate="🏎️ <b>Piloto:</b> %{customdata}<br>📍 <b>Circuito:</b> %{customdata}<br>🏁 <b>Resultado (C1/C2):</b> %{customdata}<br>🏆 <b>Puntos Acumulados:</b> %{y:.2f} pts<br>⚖️ <b>Lastre Inicial:</b> %{customdata}<extra></extra>"
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

            # 🛠️ CONTEO SEGURO DE POLES (Apuntando a las columnas H, J, L, N de puntos reales)
            for f_pole in filas_pole_hoja1:
                for piloto, col_idx in indices_pilotos_graf.items():
                    columna_puntos_idx = col_idx + 1
                    val_pole = df_hoja1_graf.iloc[f_pole, columna_puntos_idx]
                    
                    if pd.notna(val_pole) and str(val_pole).strip() != "":
                        try:
                            # Convertimos el punto cargado manualmente en tu Hoja1
                            num_puntos_pole = float(str(val_pole).replace(",", ".").strip())
                            
                            # Condición exacta: sumamos la pole si y solo si la celda tiene exactamente 1.0 punto
                            if num_puntos_pole == 1.0:
                                poles_totales_por_piloto[piloto] += 1
                        except:
                            pass


            # Extracción de posiciones promedio desde Hoja1
            # Recorremos de forma segura cada una de las fechas disputadas
            for idx_est in range(len(filas_c1)):
                for piloto, col_idx in indices_pilotos_graf.items():
                    val_c1 = df_hoja1_graf.iloc[filas_c1[idx_est], col_idx]
                    if pd.notna(val_c1):
                        try:
                            num_c1 = float(str(val_c1).upper().replace("P", "").strip())
                            if num_c1 > 0: 
                                posiciones_c1_por_piloto[piloto].append(num_c1)
                        except: 
                            pass
                    
                    if idx_est < len(filas_c2):
                        val_c2 = df_hoja1_graf.iloc[filas_c2[idx_est], col_idx]
                        if pd.notna(val_c2):
                            try:
                                num_c2 = float(str(val_c2).upper().replace("P", "").strip())
                                if num_c2 > 0: 
                                    posiciones_c2_por_piloto[piloto].append(num_c2)
                            except: 
                                pass
                    
                    if idx_est < len(filas_c2):
                        val_c2 = df_hoja1_graf.iloc[filas_c2[idx_est], col_idx]
                        if pd.notna(val_c2):
                            try:
                                num_c2 = float(str(val_c2).upper().replace("P", "").strip())
                                if num_c2 > 0: 
                                    posiciones_c2_por_piloto[piloto].append(num_c2)
                            except: 
                                pass

            # --- COMPILACIÓN DEL REPORTE FINAL ---
            reporte_tarjetas = []
            for piloto in pilotos_torneo:
                lista_c1 = posiciones_c1_por_piloto.get(piloto, [])
                lista_c2 = posiciones_c2_por_piloto.get(piloto, [])
                prom_c1 = sum(lista_c1) / len(lista_c1) if lista_c1 else 0.0
                prom_c2 = sum(lista_c2) / len(lista_c2) if lista_c2 else 0.0
                poles_reales = poles_totales_por_piloto.get(piloto, 0)
                
                total_mangas = lista_c1 + lista_c2
                prom_general = sum(total_mangas) / len(total_mangas) if total_mangas else 0.0

                reporte_tarjetas.append({
                    "Piloto": piloto, 
                    "Promedio C1": prom_c1 if prom_c1 > 0 else 99.0,
                    "Promedio C2": prom_c2 if prom_c2 > 0 else 99.0, 
                    "Promedio General": prom_general if prom_general > 0 else 99.0,
                    "Poles": poles_reales
                })

            df_reporte_tarjetas = pd.DataFrame(reporte_tarjetas).sort_values(by="Promedio General")

            # --- MAQUILLAJE DE LAS TARJETAS (ESTILO LASTRE) ---
            st.markdown("""
            <style>
            .tarjeta-simulacion-f1 {
                background-color: #161925;
                border-left: 5px solid #e10600;
                border-radius: 10px;
                padding: 18px;
                margin-bottom: 15px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.5);
            }
            .titulo-simulacion-f1 {
                margin: 0; font-size: 11px; color: #9fa6b2; text-transform: uppercase; font-weight: bold; letter-spacing: 1px;
            }
            .piloto-simulacion-f1 {
                margin: 0; padding: 4px 0; color: #ffffff; font-size: 26px; font-weight: bold;
            }
            .divisor-simulacion-f1 {
                margin: 8px 0; border-color: #2a2f45;
            }
            .bloque-valores-f1 {
                display: flex; justify-content: space-between; align-items: center;
            }
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
                p_gen = f"P{row['Promedio General']:.1f}" if row["Promedio General"] != 99.0 else "-"
                p_c1 = f"P{row['Promedio C1']:.1f}" if row["Promedio C1"] != 99.0 else "-"
                p_c2 = f"P{row['Promedio C2']:.1f}" if row["Promedio C2"] != 99.0 else "-"
                p_poles = int(row["Poles"])
                
                color_borde_pista = "#00e676" if idx_c == 0 else "#e10600"
                
                with cols_grid[idx_c % 4]:
                    st.markdown(f"""
                        <div class="tarjeta-simulacion-f1" style="border-left: 5px solid {color_borde_pista};">
                            <p class="titulo-simulacion-f1">Rendimiento en Pista</p>
                            <h3 class="piloto-simulacion-f1">{p_name}</h3>
                            <hr class="divisor-simulacion-f1">
                            <div class="bloque-valores-f1">
                                <div class="sub-metrica-f1">
                                    <p class="texto-gris-f1">Ritmo General:</p>
                                    <p class="texto-blanco-bold-f1" style="color: #00e676;">{p_gen}</p>
                                </div>
                                <div class="sub-metrica-der-f1">
                                    <p class="texto-gris-f1">Poles / C1 / C2:</p>
                                    <p class="texto-rojo-bold-f1" style="color: #00b0ff;">🥇 {p_poles} <span style="color:#ffffff; font-size:11px; font-weight:normal;">({p_c1}/{p_c2})</span></p>
                                </div>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)

            st.write("#### 📊 Tabla de Resumen Detallada")
            st.dataframe(
                df_reporte_tarjetas.assign(
                    **{
                        "Promedio General": lambda x: x["Promedio General"].apply(lambda v: f"P{v:.1f}" if v != 99.0 else "-"),
                        "Promedio C1": lambda x: x["Promedio C1"].apply(lambda v: f"P{v:.1f}" if v != 99.0 else "-"),
                        "Promedio C2": lambda x: x["Promedio C2"].apply(lambda v: f"P{v:.1f}" if v != 99.0 else "-"),
                        "Poles": lambda x: x["Poles"].apply(lambda v: f"🥇 {int(v)} Poles")
                    }
                ),
                use_container_width=True, hide_index=True
            )

            rey_pole = max(poles_totales_por_piloto, key=poles_totales_por_piloto.get)
            max_poles = poles_totales_por_piloto[rey_pole]
            if max_poles > 0:
                st.success(f"⏱️ **Rey de los Sábados:** El piloto con más Pole Positions es **{rey_pole}** con un total de **{max_poles} Poles**.")

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

        # --- SECCIÓN CSS UNIFICADA PARA LAS NUEVAS TARJETAS ---
        st.markdown("""
        <style>
        .columna-tiempos-box {
            background-color: #0f111a;
            border: 1px solid #2a2f45;
            border-radius: 12px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .tarjeta-tiempo-f1 {
            background-color: #161925;
            border-left: 5px solid #e10600;
            border-radius: 8px;
            padding: 10px 14px;
            margin-bottom: 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 4px 8px rgba(0,0,0,0.3);
        }
        .info-izq-tiempo {
            display: flex;
            flex-direction: column;
        }
        .piloto-tiempo-f1 {
            color: #ffffff;
            font-size: 15px;
            font-weight: bold;
        }
        .tiempo-vuelta-f1 {
            color: #00ff88;
            font-size: 13px;
            font-family: monospace;
            margin-top: 1px;
        }
        .gap-der-tiempo {
            text-align: right;
            font-size: 14px;
            font-weight: bold;
            font-family: monospace;
        }
        </style>
        """, unsafe_allow_html=True)

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
    st.write("Consulta el peso acumulado en pista y el rendimiento de los pilotos con lastre.")

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

            # -----------------------------------------------------------------
            # 📋 ACÁ ABAJO CONTINÚA TU LÓGICA ORIGINAL DE LAS TARJETAS DE KILOS
            # -----------------------------------------------------------------
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

            # =========================================================================
            # 🏋️ MAPA VISUAL DE EFECTIVIDAD POR PESO (ESTILO F1 TV - SIN TABLAS)
            # =========================================================================
            st.markdown("---")
            st.subheader("🏋️ Análisis de Rendimiento con Lastre Técnico")
            st.write("Evaluación histórica del promedio de puntos sumados y la cantidad de fechas disputadas en cada rango de peso:")

            rangos_peso = ["0-20 Kg", "21-40 Kg", "41-60 Kg", "61-80 Kg+"]
            historial_peso_puntos = {p: {r: [] for r in rangos_peso} for p in pilotos_torneo}

            for idx_c, fila_total in enumerate(filas_lastre_fecha_real[:ultima_fecha_real_num]):
                for piloto, col_idx in indices_pilotos_hoja1.items():
                    val_puntos = df_hoja1_sim.iloc[fila_total, col_idx]
                    puntos_limpios = str(val_puntos).replace(',', '.', 1) if pd.notna(val_puntos) else "0.0"
                    pts_fecha = float(puntos_limpios) if puntos_limpios.replace('.','',1).isdigit() else 0.0
                    
                    kilos_largada = 0.0
                    if idx_c > 0:
                        idx_bloque_anterior = idx_c - 1
                        if idx_bloque_anterior < len(filas_lastre_acum_real):
                            fila_lastre_anterior = filas_lastre_acum_real[idx_bloque_anterior]
                            val_lastre = df_hoja1_sim.iloc[fila_lastre_anterior, col_idx]
                            if pd.notna(val_lastre):
                                kilos_largada = float(str(val_lastre).upper().replace("KG","").strip())

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

            st.info("💡 **Análisis de Telemetría:** Las tarjetas ordenan automáticamente a los pilotos de mayor a menor efectividad dentro de cada rango de plomo.")

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
            ultima_fecha_real_num = 7

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
        # 🧮 CALCULADORA MATEMÁTICA DE TÍTULO ACTUALIZADA AL 75% (46.75 PTS MAX)
        # =========================================================================
        st.markdown("---")
        st.subheader("🧮 Calculadora de Título Matemática")

        fecha_final_simulada_num = ultima_fecha_real_num + cant_fechas_a_proyectar
        fechas_restantes_campeonato = max(0, 10 - fecha_final_simulada_num)
        
        # 🛠️ REGLAMENTO OFICIAL ACTUALIZADO: 25 (C1) + 18.75 (C2) + 1 (Pole) + 1 (VR C1) + 1 (VR C2) = 46.75 pts
        puntos_maximos_por_fecha = 46.75
        puntos_en_juego_totales = fechas_restantes_campeonato * puntos_maximos_por_fecha

        if fechas_restantes_campeonato > 0:
            st.write(f"📊 Al finalizar la simulación, quedarán **{fechas_restantes_campeonato} fechas en juego** (Máximo absoluto disponible: **{puntos_en_juego_totales:.1f} pts**).")
            
            pilotos_con_chances = []
            pilotos_eliminados = []
            
            for _, row_p in df_proyeccion_final.iterrows():
                p_nombre = row_p["Piloto"]
                p_puntos = row_p["Puntos Finales Proyectados"]
                distancia_al_lider = puntos_lider_proy - p_puntos
                
                if distancia_al_lider <= puntos_en_juego_totales:
                    if p_nombre == lider_proyectado:
                        pilotos_con_chances.append(f"👑 **{p_nombre}** (Líder actual)")
                    else:
                        pilotos_con_chances.append(f"🏎️ **{p_nombre}** (A {distancia_al_lider:.2f} pts del líder)")
                else:
                    pilotos_eliminados.append(p_nombre)
            
            col_vivos, col_eliminados = st.columns(2)
            with col_vivos:
                st.markdown("🟢 **Siguen en la Pelea Matemática:**")
                for p_vivo in pilotos_con_chances: st.write(p_vivo)
            with col_eliminados:
                st.markdown("🔴 **Matemáticamente Sin Chances:**")
                if pilotos_eliminados:
                    for p_chau in pilotos_eliminados: st.write(f"❌ {p_chau}")
                else:
                    st.write("¡Ninguno! Todos los pilotos mantienen chances matemáticas.")
        else:
            st.balloons()
            st.markdown(f"## 🏆 ¡TENEMOS CAMPEÓN DEL TORNEO! ##\nMatemáticamente, **{lider_proyectado}** se consagra Campeón Oficial del Torneo TC2000 con **{puntos_lider_proy:.2f} puntos**.")

        st.markdown("---")
        # =========================================================================
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
            min-width: 35px;
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
