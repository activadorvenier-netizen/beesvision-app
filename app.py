import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
import os
import io
import time

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA
# ------------------------------------------------

st.set_page_config(
    page_title="Coaching Ventas",
    page_icon="📈",
    layout="wide"
)

if "form_id" not in st.session_state:
    st.session_state.form_id = 0

# ------------------------------------------------
# CONFIGURACIÓN GOOGLE SHEETS
# ------------------------------------------------

SHEET_ID = "1qRsvFn62DlMYx4xHtbboYNiOh2flj-JQcaiJ9taTmjE"

# =========================================================
# NOMBRES CORRECTOS DE LAS HOJAS (EN MAYÚSCULAS)
# =========================================================
HOJAS = {
    "auditores": "AUDITOR",
    "auditados": "AUDITADO",
    "empresas": "EMPRESAS",
    "localidades": "LOCALIDADES",
    "notas": "NOTAS",
    "pilares": "PILARES",
    "preguntas": "PREGUNTAS",
    "respuestas": "RESPUESTAS",
    "resumen": "RESUMEN"
}
# =========================================================

def conectar_gsheets():
    """Conecta a Google Sheets usando Service Account"""
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        
        # USAR SOLO VARIABLES DE ENTORNO
        import json
        
        # Construir el diccionario con las variables de entorno
        creds_dict = {
            "type": os.getenv("GCP_TYPE", "service_account"),
            "project_id": os.getenv("GCP_PROJECT_ID", ""),
            "private_key_id": os.getenv("GCP_PRIVATE_KEY_ID", ""),
            "private_key": os.getenv("GCP_PRIVATE_KEY", "").replace("\\n", "\n"),
            "client_email": os.getenv("GCP_CLIENT_EMAIL", ""),
            "client_id": os.getenv("GCP_CLIENT_ID", ""),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_x509_cert_url": os.getenv("GCP_CLIENT_X509_CERT_URL", "")
        }
        
        # Verificar que la private_key no esté vacía
        if not creds_dict["private_key"] or not creds_dict["client_email"]:
            st.error("❌ Error: Credenciales incompletas. Verifica las variables de entorno en Render.")
            st.info("Variables requeridas: GCP_PRIVATE_KEY, GCP_CLIENT_EMAIL, GCP_PROJECT_ID")
            return None
        
        # Crear credenciales
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open_by_key(SHEET_ID)
    except Exception as e:
        st.error(f"Error conectando a Google Sheets: {e}")
        st.info("Verifica que las variables de entorno en Render estén correctamente configuradas.")
        return None

def leer_hoja(hoja_nombre):
    """Lee una hoja de Google Sheets y la convierte en DataFrame"""
    try:
        sheet = conectar_gsheets()
        if sheet is None:
            return pd.DataFrame()
        
        worksheet = sheet.worksheet(hoja_nombre)
        data = worksheet.get_all_values()
        
        if len(data) == 0:
            return pd.DataFrame()
        
        headers = data[0]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=headers)
        return df
    except Exception as e:
        print(f"Error leyendo hoja {hoja_nombre}: {e}")
        return pd.DataFrame()

def guardar_hoja(df, hoja_nombre):
    """Guarda un DataFrame en Google Sheets (reemplaza toda la hoja)"""
    try:
        sheet = conectar_gsheets()
        if sheet is None:
            return False
        
        worksheet = sheet.worksheet(hoja_nombre)
        worksheet.clear()
        
        if df.empty:
            worksheet.update([df.columns.tolist()])
        else:
            datos = [df.columns.tolist()] + df.values.tolist()
            worksheet.update(datos)
        
        return True
    except Exception as e:
        print(f"Error guardando hoja {hoja_nombre}: {e}")
        return False

def agregar_fila_hoja(df, hoja_nombre):
    """Agrega filas a una hoja (append)"""
    try:
        sheet = conectar_gsheets()
        if sheet is None:
            return False
        
        worksheet = sheet.worksheet(hoja_nombre)
        
        existing = worksheet.get_all_values()
        if len(existing) == 0:
            worksheet.append_row(df.columns.tolist())
        
        for _, row in df.iterrows():
            worksheet.append_row(row.astype(str).tolist())
        
        return True
    except Exception as e:
        print(f"Error agregando filas a {hoja_nombre}: {e}")
        return False

# ------------------------------------------------
# CARGA DE DATOS (con caché para rendimiento)
# ------------------------------------------------

@st.cache_data(ttl=300)
def cargar_datos():
    """Carga todos los datos de Google Sheets"""
    return {
        "auditores": leer_hoja(HOJAS["auditores"]),
        "auditados": leer_hoja(HOJAS["auditados"]),
        "empresas": leer_hoja(HOJAS["empresas"]),
        "localidades": leer_hoja(HOJAS["localidades"]),
        "notas": leer_hoja(HOJAS["notas"]),
        "pilares": leer_hoja(HOJAS["pilares"]),
        "preguntas": leer_hoja(HOJAS["preguntas"]),
        "respuestas": leer_hoja(HOJAS["respuestas"]),
        "resumen": leer_hoja(HOJAS["resumen"])
    }

# ------------------------------------------------
# FUNCIONES DE NEGOCIO
# ------------------------------------------------

def obtener_localidades(empresa, df_localidades):
    if df_localidades.empty:
        return []
    if empresa:
        return df_localidades[df_localidades["Empresas"] == empresa]["Localidades"].dropna().astype(str).tolist()
    return df_localidades["Localidades"].dropna().astype(str).tolist()

def obtener_auditados(empresa, auditor, puesto_auditor, df_auditados):
    """Obtiene auditados filtrados por empresa y según el puesto del auditor"""
    if df_auditados.empty:
        return []
    df_filtrado = df_auditados.copy()
    
    if empresa:
        df_filtrado = df_filtrado[df_filtrado["Empresa"] == empresa]
    
    if puesto_auditor == "Jefe de Ventas":
        df_filtrado = df_filtrado[df_filtrado["Puesto"] == "Supervisor"]
    elif puesto_auditor == "Supervisor":
        df_filtrado = df_filtrado[df_filtrado["Puesto"].isin(["Promotor", "Merchand"])]
        if auditor:
            df_filtrado = df_filtrado[df_filtrado["Auditor"] == auditor]
    
    return df_filtrado["Auditado"].dropna().astype(str).tolist()

def obtener_preguntas(puesto, df_preguntas):
    if df_preguntas.empty:
        return pd.DataFrame()
    return df_preguntas[df_preguntas["Puesto"] == puesto]

def obtener_categoria_mejorar(df_preguntas_filtrado, respuestas, df_notas):
    if df_preguntas_filtrado.empty or not respuestas:
        return ""
    
    categorias = {}
    for _, row in df_preguntas_filtrado.iterrows():
        pregunta = row["Pregunta"]
        categoria = row["Categoría"]
        puntos_max = float(row["Puntos Máximo"])
        
        if pregunta in respuestas:
            nota = respuestas[pregunta]
            if nota:
                peso_row = df_notas[df_notas["Nota"] == nota]
                if not peso_row.empty:
                    peso = float(peso_row["Peso"].iloc[0])
                    puntaje = peso * puntos_max
                    if categoria not in categorias:
                        categorias[categoria] = {"total": 0, "count": 0}
                    categorias[categoria]["total"] += puntaje
                    categorias[categoria]["count"] += 1
    
    promedios = {}
    for cat, datos in categorias.items():
        promedios[cat] = datos["total"] / datos["count"] if datos["count"] > 0 else 0
    
    if promedios:
        peor_categoria = min(promedios, key=promedios.get)
        return peor_categoria
    return ""

# ------------------------------------------------
# SIDEBAR
# ------------------------------------------------

st.sidebar.image(
    "Logo Grupo Venier.png",
    use_container_width=True
)

with st.sidebar:
    seleccion = st.radio(
        "Navegación",
        [
            "📝 Nuevo Coaching",
            "📊 Dashboard",
            "🎯 Categorías a Mejorar",
            "📂 Historial",
            "⚙️ Maestros"
        ],
        index=0
    )
    
    st.markdown("---")
    
    # Botón para recargar datos - CORREGIDO SIN st.rerun()
    if st.button("🔄 Recargar Datos", use_container_width=True):
        st.cache_data.clear()
        st.success("✅ Caché limpiada. La página se recargará automáticamente.")
        time.sleep(1)
        st.experimental_rerun()

    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: gray; font-size: 12px; padding: 10px;'>
            By Pato Frangi
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================================================
# NUEVO COACHING
# =========================================================

if seleccion == "📝 Nuevo Coaching":
    st.title("📝 Nuevo Coaching")

    # Cargar datos
    datos = cargar_datos()
    df_auditores = datos["auditores"]
    df_auditados = datos["auditados"]
    df_empresas = datos["empresas"]
    df_localidades = datos["localidades"]
    df_notas = datos["notas"]
    df_preguntas = datos["preguntas"]

    # Listas para selectores
    lista_auditores = df_auditores["Auditor"].dropna().astype(str).tolist() if not df_auditores.empty else []
    lista_empresas = df_empresas["Empresas"].dropna().astype(str).tolist() if not df_empresas.empty else []
    lista_notas = df_notas["Nota"].dropna().astype(str).tolist() if not df_notas.empty else []

    # ------------------------------------------------
    # INICIALIZAR SESSION STATE
    # ------------------------------------------------
    
    if "selected_auditor" not in st.session_state:
        st.session_state.selected_auditor = ""
    if "selected_empresa" not in st.session_state:
        st.session_state.selected_empresa = ""
    if "selected_auditado" not in st.session_state:
        st.session_state.selected_auditado = ""
    if "selected_localidad" not in st.session_state:
        st.session_state.selected_localidad = ""

    # ------------------------------------------------
    # SELECTORES (FUERA DEL FORMULARIO) - SIN st.rerun()
    # ------------------------------------------------
    
    col1, col2 = st.columns(2)
    
    with col1:
        fecha = st.date_input("📅 Fecha", value=date.today())
        
        # Selector de Auditor
        auditor_idx = 0
        if st.session_state.selected_auditor in lista_auditores:
            auditor_idx = lista_auditores.index(st.session_state.selected_auditor) + 1
        
        auditor = st.selectbox(
            "👤 Auditor",
            [""] + lista_auditores,
            index=auditor_idx,
            key="auditor_select"
        )
        
        if auditor != st.session_state.selected_auditor:
            st.session_state.selected_auditor = auditor
            st.session_state.selected_auditado = ""
            st.session_state.selected_localidad = ""
        
        # Selector de Empresa
        empresa_idx = 0
        if st.session_state.selected_empresa in lista_empresas:
            empresa_idx = lista_empresas.index(st.session_state.selected_empresa) + 1
        
        empresa = st.selectbox(
            "🏢 Empresa",
            [""] + lista_empresas,
            index=empresa_idx,
            key="empresa_select"
        )
        
        if empresa != st.session_state.selected_empresa:
            st.session_state.selected_empresa = empresa
            st.session_state.selected_auditado = ""
            st.session_state.selected_localidad = ""
    
    with col2:
        # Obtener puesto del auditor seleccionado
        puesto_auditor = ""
        if st.session_state.selected_auditor and not df_auditores.empty:
            puesto_row = df_auditores[df_auditores["Auditor"] == st.session_state.selected_auditor]
            if not puesto_row.empty:
                puesto_auditor = puesto_row["Puesto"].iloc[0]
        
        # Obtener auditados disponibles según filtros
        auditados_disponibles = obtener_auditados(
            st.session_state.selected_empresa,
            st.session_state.selected_auditor,
            puesto_auditor,
            df_auditados
        )
        
        # Selector de Auditado
        auditado_idx = 0
        if st.session_state.selected_auditado in auditados_disponibles:
            auditado_idx = auditados_disponibles.index(st.session_state.selected_auditado) + 1
        
        auditado = st.selectbox(
            "👤 Auditado",
            [""] + auditados_disponibles,
            index=auditado_idx,
            key="auditado_select"
        )
        
        if auditado != st.session_state.selected_auditado:
            st.session_state.selected_auditado = auditado
        
        # Obtener localidades disponibles según empresa
        localidades_disponibles = obtener_localidades(
            st.session_state.selected_empresa,
            df_localidades
        )
        
        # Selector de Localidad
        localidad_idx = 0
        if st.session_state.selected_localidad in localidades_disponibles:
            localidad_idx = localidades_disponibles.index(st.session_state.selected_localidad) + 1
        
        localidad = st.selectbox(
            "📍 Localidad",
            [""] + localidades_disponibles,
            index=localidad_idx,
            key="localidad_select"
        )
        
        if localidad != st.session_state.selected_localidad:
            st.session_state.selected_localidad = localidad

    st.divider()

    # ------------------------------------------------
    # PREGUNTAS
    # ------------------------------------------------

    st.subheader("📋 Preguntas")

    auditado_seleccionado = st.session_state.selected_auditado

    if auditado_seleccionado:
        # Obtener puesto del auditado
        puesto_auditado = ""
        if not df_auditados.empty:
            puesto_row = df_auditados[df_auditados["Auditado"] == auditado_seleccionado]
            if not puesto_row.empty:
                puesto_auditado = puesto_row["Puesto"].iloc[0]

        if not puesto_auditado:
            st.warning("⚠️ No se encontró el puesto del auditado. Verifique en Maestros.")
        else:
            # Obtener preguntas para el puesto
            df_preguntas_filtrado = obtener_preguntas(puesto_auditado, df_preguntas)

            if df_preguntas_filtrado.empty:
                st.warning(f"⚠️ No hay preguntas para el puesto '{puesto_auditado}'")
            else:
                st.info(f"📌 Puesto del auditado: **{puesto_auditado}** | 📋 Preguntas encontradas: **{len(df_preguntas_filtrado)}**")

                # ------------------------------------------------
                # FORMULARIO DE PREGUNTAS Y GUARDADO
                # ------------------------------------------------
                
                with st.form(key=f"coaching_form_{st.session_state.form_id}"):
                    respuestas = {}

                    for idx, (_, row) in enumerate(df_preguntas_filtrado.iterrows(), 1):
                        pregunta = row["Pregunta"]
                        puntos_maximo = float(row["Puntos Máximo"])
                        categoria = row["Categoría"]
                        pregunta_id = row["ID"]

                        st.markdown(f"**{idx}. {pregunta}**")
                        st.caption(f"⚖️ Puntos Máximo: {puntos_maximo} | 📂 Categoría: {categoria}")

                        nota = st.radio(
                            label="",
                            options=[""] + lista_notas,
                            key=f"nota_{pregunta_id}_{st.session_state.form_id}",
                            horizontal=True,
                            label_visibility="collapsed"
                        )

                        if nota:
                            respuestas[pregunta] = nota

                        st.divider()

                    # Botón Guardar
                    submitted = st.form_submit_button("💾 Guardar Coaching", use_container_width=True, type="primary")

                # ------------------------------------------------
                # PROCESAR GUARDADO
                # ------------------------------------------------
                
                if submitted:
                    # Validar campos obligatorios
                    if not st.session_state.selected_auditor or not st.session_state.selected_auditado or not st.session_state.selected_empresa or not st.session_state.selected_localidad:
                        st.error("⚠️ Complete todos los campos obligatorios antes de guardar.")
                        st.stop()

                    # Verificar que todas las preguntas tengan respuesta
                    preguntas_sin_respuesta = []
                    for _, row in df_preguntas_filtrado.iterrows():
                        pregunta = row["Pregunta"]
                        if pregunta not in respuestas or not respuestas[pregunta]:
                            preguntas_sin_respuesta.append(pregunta)

                    if preguntas_sin_respuesta:
                        st.error(f"⚠️ Faltan responder {len(preguntas_sin_respuesta)} preguntas.")
                        for p in preguntas_sin_respuesta[:5]:
                            st.warning(f"❓ {p}")
                        if len(preguntas_sin_respuesta) > 5:
                            st.warning(f"... y {len(preguntas_sin_respuesta) - 5} más")
                        st.stop()

                    # Mostrar spinner de carga
                    with st.spinner("⏳ Guardando coaching, por favor espere..."):
                        time.sleep(1)

                        # Preparar datos para guardar
                        datos_coaching = []
                        total_puntaje = 0
                        total_maximo = 0
                        
                        categorias_por_nombre = {}

                        for _, row in df_preguntas_filtrado.iterrows():
                            pregunta = row["Pregunta"]
                            puntos_maximo = float(row["Puntos Máximo"])
                            nota = respuestas.get(pregunta, "")
                            pilar = row["Pilar"]
                            categoria = row["Categoría"]

                            peso = 0
                            puntaje_final = 0
                            if nota and not df_notas.empty:
                                peso_row = df_notas[df_notas["Nota"] == nota]
                                if not peso_row.empty:
                                    peso = float(peso_row["Peso"].iloc[0])
                                    puntaje_final = peso * puntos_maximo

                            total_puntaje += puntaje_final
                            total_maximo += puntos_maximo

                            if categoria not in categorias_por_nombre:
                                categorias_por_nombre[categoria] = {"total": 0, "count": 0}
                            categorias_por_nombre[categoria]["total"] += puntaje_final
                            categorias_por_nombre[categoria]["count"] += 1

                            datos_coaching.append({
                                "Fecha": str(fecha),
                                "Auditor": st.session_state.selected_auditor,
                                "Auditado": st.session_state.selected_auditado,
                                "Empresa": st.session_state.selected_empresa,
                                "Localidad": st.session_state.selected_localidad,
                                "Pilar": pilar,
                                "Pregunta": pregunta,
                                "Nota": nota,
                                "Peso": peso,
                                "Puntos Máximo": puntos_maximo,
                                "Puntaje Final": puntaje_final,
                                "Categoría": categoria
                            })

                        df_guardado = pd.DataFrame(datos_coaching)

                        # Calcular score
                        score = (total_puntaje / total_maximo * 100) if total_maximo > 0 else 0
                        
                        categorias_mejorar = []
                        
                        if score < 100:
                            promedios_categorias = {}
                            for categoria, datos in categorias_por_nombre.items():
                                promedio = datos["total"] / datos["count"] if datos["count"] > 0 else 0
                                rows_categoria = df_preguntas_filtrado[df_preguntas_filtrado["Categoría"] == categoria]
                                if not rows_categoria.empty:
                                    puntos_max_categoria = float(rows_categoria.iloc[0]["Puntos Máximo"])
                                    porcentaje = (promedio / puntos_max_categoria) * 100
                                    promedios_categorias[categoria] = porcentaje
                            
                            if promedios_categorias:
                                min_porcentaje = min(promedios_categorias.values())
                                categorias_mejorar = [cat for cat, pct in promedios_categorias.items() if pct == min_porcentaje]

                        # =========================================================
                        # GUARDAR EN RESPUESTAS
                        # =========================================================
                        
                        if agregar_fila_hoja(df_guardado, "RESPUESTAS"):
                            st.success("✅ Datos guardados en RESPUESTAS")
                            
                            # =========================================================
                            # GUARDAR EN RESUMEN
                            # =========================================================
                            
                            categorias_str = ", ".join(categorias_mejorar) if categorias_mejorar else "Ninguna"
                            df_resumen = pd.DataFrame([{
                                "Fecha": str(fecha),
                                "Auditor": st.session_state.selected_auditor,
                                "Auditado": st.session_state.selected_auditado,
                                "Empresa": st.session_state.selected_empresa,
                                "Localidad": st.session_state.selected_localidad,
                                "Puntaje Final": round(score, 1),
                                "Categorías a Mejorar": categorias_str
                            }])
                            
                            if agregar_fila_hoja(df_resumen, "RESUMEN"):
                                st.success("✅ Datos guardados en RESUMEN")
                            else:
                                st.warning("⚠️ Error guardando en RESUMEN")
                            
                            # =========================================================
                            # ACTUALIZAR CONTROL_COACHINGS
                            # =========================================================
                            
                            try:
                                bimestre_num = None
                                mes = fecha.month
                                if mes in [1, 2]:
                                    bimestre_num = 1
                                elif mes in [3, 4]:
                                    bimestre_num = 2
                                elif mes in [5, 6]:
                                    bimestre_num = 3
                                elif mes in [7, 8]:
                                    bimestre_num = 4
                                elif mes in [9, 10]:
                                    bimestre_num = 5
                                elif mes in [11, 12]:
                                    bimestre_num = 6
                                
                                if bimestre_num:
                                    sheet = conectar_gsheets()
                                    if sheet:
                                        try:
                                            ws_control = sheet.worksheet("CONTROL_COACHINGS")
                                        except:
                                            ws_control = sheet.add_worksheet("CONTROL_COACHINGS", rows=100, cols=10)
                                            ws_control.update([["Auditado", "Coaching 1", "Coaching 2", "Coaching 3", "Coaching 4", "Coaching 5", "Coaching 6"]])
                                        
                                        datos_control = ws_control.get_all_values()
                                        fila_auditado = None
                                        
                                        for i, row in enumerate(datos_control):
                                            if len(row) > 0 and row[0] == st.session_state.selected_auditado:
                                                fila_auditado = i + 1
                                                break
                                        
                                        if fila_auditado is None:
                                            nueva_fila = [st.session_state.selected_auditado] + [""] * 6
                                            ws_control.append_row(nueva_fila)
                                            fila_auditado = len(datos_control) + 1
                                        
                                        columna = bimestre_num + 1
                                        ws_control.update_cell(fila_auditado, columna, str(fecha))
                                        
                                        st.success(f"✅ CONTROL_COACHINGS actualizado: {st.session_state.selected_auditado} - Coaching {bimestre_num} = {fecha}")
                                        
                            except Exception as e:
                                st.error(f"❌ Error en CONTROL_COACHINGS: {e}")
                            
                            # =========================================================
                            # ENVIAR EMAIL AL AUDITADO
                            # =========================================================
                            
                            try:
                                email_auditado = None
                                if not df_auditados.empty and "Email" in df_auditados.columns:
                                    email_row = df_auditados[df_auditados["Auditado"] == st.session_state.selected_auditado]
                                    if not email_row.empty:
                                        email_auditado = email_row["Email"].iloc[0]
                                
                                if email_auditado:
                                    categorias_texto = ", ".join(categorias_mejorar) if categorias_mejorar else "No hay categorías a mejorar"
                                    st.info(f"📧 Email a {email_auditado}: Score {score:.1f}% - {categorias_texto}")
                                else:
                                    st.warning(f"⚠️ No se encontró email para el auditado: {st.session_state.selected_auditado}")
                                    
                            except Exception as e:
                                st.warning(f"⚠️ Error en email: {e}")

                            # Limpiar caché
                            st.cache_data.clear()

                            # Mostrar mensaje de éxito
                            if score >= 85:
                                st.success(f"🏆 ¡Excelente! Score: {score:.1f}%")
                                st.balloons()
                            elif score >= 70:
                                st.success(f"📊 Buen desempeño. Score: {score:.1f}%")
                            else:
                                st.warning(f"📈 Área de oportunidad. Score: {score:.1f}%")

                            if categorias_mejorar:
                                st.info(f"🎯 Categorías a Mejorar: **{categorias_str}**")
                            else:
                                st.success("🎯 ¡No hay categorías a mejorar!")

                            # Reiniciar formulario - SIN rerun
                            st.session_state.form_id += 1
                            st.session_state.selected_auditor = ""
                            st.session_state.selected_empresa = ""
                            st.session_state.selected_auditado = ""
                            st.session_state.selected_localidad = ""
                            
                            # NO usar rerun aquí - el usuario puede seguir navegando
                            
                        else:
                            st.error("❌ Error guardando en RESPUESTAS")
    else:
        st.info("👆 Seleccione un Auditado para ver las preguntas.")

# =========================================================
# DASHBOARD
# =========================================================

elif seleccion == "📊 Dashboard":
    st.title("📊 Dashboard Coaching Ventas")

    # Cargar datos
    datos = cargar_datos()
    df_respuestas = datos["respuestas"]

    if df_respuestas.empty:
        st.warning("⚠️ No existen sesiones de coaching cargadas")
        st.stop()

    # Convertir tipos
    df_respuestas["Fecha"] = pd.to_datetime(df_respuestas["Fecha"], errors="coerce")
    df_respuestas["Puntaje Final"] = pd.to_numeric(df_respuestas["Puntaje Final"], errors="coerce")
    df_respuestas["Puntos Máximo"] = pd.to_numeric(df_respuestas["Puntos Máximo"], errors="coerce")
    df_respuestas["% Cumplimiento"] = (df_respuestas["Puntaje Final"] / df_respuestas["Puntos Máximo"] * 100).round(2)

    # =========================================================
    # DEFINIR BIMESTRES
    # =========================================================
    
    def obtener_bimestre(fecha):
        """Devuelve el número de bimestre (1-6) según la fecha"""
        if pd.isna(fecha):
            return None
        mes = fecha.month
        if mes in [1, 2]:
            return 1
        elif mes in [3, 4]:
            return 2
        elif mes in [5, 6]:
            return 3
        elif mes in [7, 8]:
            return 4
        elif mes in [9, 10]:
            return 5
        elif mes in [11, 12]:
            return 6
        return None

    def obtener_nombre_bimestre(bimestre):
        """Devuelve el nombre del bimestre"""
        nombres = {
            1: "Coaching 1 (Ene-Feb)",
            2: "Coaching 2 (Mar-Abr)",
            3: "Coaching 3 (May-Jun)",
            4: "Coaching 4 (Jul-Ago)",
            5: "Coaching 5 (Sep-Oct)",
            6: "Coaching 6 (Nov-Dic)"
        }
        return nombres.get(bimestre, f"Coaching {bimestre}")

    # Agregar columna de bimestre
    df_respuestas["Bimestre"] = df_respuestas["Fecha"].apply(obtener_bimestre)
    df_respuestas["Bimestre_Nombre"] = df_respuestas["Bimestre"].apply(obtener_nombre_bimestre)

    # =========================================================
    # FILTROS
    # =========================================================

    st.subheader("🔎 Filtros")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        auditores_dash = ["Todos"] + sorted(df_respuestas["Auditor"].dropna().unique().tolist())
        auditor_filtro = st.selectbox("👤 Auditor", auditores_dash, key="dash_auditor")

    with col2:
        empresas_dash = ["Todos"] + sorted(df_respuestas["Empresa"].dropna().unique().tolist())
        empresa_filtro = st.selectbox("🏢 Empresa", empresas_dash, key="dash_empresa")

    with col3:
        auditados_dash = ["Todos"] + sorted(df_respuestas["Auditado"].dropna().unique().tolist())
        auditado_filtro = st.selectbox("👤 Auditado", auditados_dash, key="dash_auditado")

    with col4:
        bimestres_dash = ["Todos"] + sorted(df_respuestas["Bimestre_Nombre"].dropna().unique().tolist())
        bimestre_filtro = st.selectbox("📅 Bimestre", bimestres_dash, key="dash_bimestre")

    # =========================================================
    # APLICAR FILTROS
    # =========================================================

    df_filtrado = df_respuestas.copy()

    if auditor_filtro != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Auditor"] == auditor_filtro]

    if empresa_filtro != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Empresa"] == empresa_filtro]

    if auditado_filtro != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Auditado"] == auditado_filtro]

    if bimestre_filtro != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Bimestre_Nombre"] == bimestre_filtro]

    if df_filtrado.empty:
        st.info("No hay datos con los filtros seleccionados")
        st.stop()

    # =========================================================
    # CÁLCULO DE SCORES - INCLUYE score_preguntas
    # =========================================================

    # Score general (todos los pilares) - SUMA total / SUMA máxima
    total_puntaje_general = df_filtrado["Puntaje Final"].sum()
    total_maximo_general = df_filtrado["Puntos Máximo"].sum()
    score_general = (total_puntaje_general / total_maximo_general * 100) if total_maximo_general > 0 else 0

    # Score por Pilar - SUMA de puntajes / SUMA de puntos máximos
    score_pilar = df_filtrado.groupby("Pilar", as_index=False).agg({
        "Puntaje Final": "sum",
        "Puntos Máximo": "sum"
    })
    score_pilar["Score"] = (score_pilar["Puntaje Final"] / score_pilar["Puntos Máximo"] * 100).round(2)
    score_pilar = score_pilar.sort_values("Score", ascending=False)

    # Score por Auditado - SUMA de puntajes / SUMA de puntos máximos
    score_auditado = df_filtrado.groupby("Auditado", as_index=False).agg({
        "Puntaje Final": "sum",
        "Puntos Máximo": "sum"
    })
    score_auditado["Score"] = (score_auditado["Puntaje Final"] / score_auditado["Puntos Máximo"] * 100).round(2)
    score_auditado = score_auditado.sort_values("Score", ascending=False)

    # Score por Pregunta - SUMA de puntajes / SUMA de puntos máximos
    score_preguntas = df_filtrado.groupby("Pregunta", as_index=False).agg({
        "Puntaje Final": "sum",
        "Puntos Máximo": "sum"
    })
    score_preguntas["Score"] = (score_preguntas["Puntaje Final"] / score_preguntas["Puntos Máximo"] * 100).round(2)
    score_preguntas = score_preguntas.sort_values("Score", ascending=False)

        # =========================================================
    # KPIs PRINCIPALES
    # =========================================================

    st.subheader("📌 Indicadores Generales")

    total_auditados = df_filtrado["Auditado"].nunique()
    total_preguntas = len(df_filtrado)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("🎯 Score General", f"{score_general:.1f}%")
    with col2:
        st.metric("👤 Auditados", total_auditados)
    with col3:
        st.metric("📝 Preguntas", total_preguntas)

    # =========================================================
    # TOP 3 MEJORES Y PEORES AUDITADOS
    # =========================================================

    st.subheader("👤 Top 3 Mejores y Peores Auditados")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 🟢 Mejores Auditados")
        for _, row in score_auditado.head(3).iterrows():
            st.markdown(
                f"""
                <div style="padding:12px;border-radius:10px;background-color:#f0fdf4;color:#000000;margin-bottom:8px;border:1px solid #86efac;">
                    <b>👤 {row['Auditado']}</b><br>
                    🎯 Score: <b>{row['Score']:.1f}%</b>
                </div>
                """,
                unsafe_allow_html=True
            )

    with col2:
        st.markdown("### 🔴 Peores Auditados")
        for _, row in score_auditado.tail(3).iterrows():
            st.markdown(
                f"""
                <div style="padding:12px;border-radius:10px;background-color:#fef2f2;color:#000000;margin-bottom:8px;border:1px solid #fca5a5;">
                    <b>👤 {row['Auditado']}</b><br>
                    🎯 Score: <b>{row['Score']:.1f}%</b>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.divider()

    # =========================================================
    # TOP 3 MEJORES Y PEORES PILARES
    # =========================================================

    st.subheader("🏆 Top 3 Mejores y Peores Pilares")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 🟢 Mejores Pilares")
        for _, row in score_pilar.head(3).iterrows():
            st.markdown(
                f"""
                <div style="padding:12px;border-radius:10px;background-color:#f0fdf4;color:#000000;margin-bottom:8px;border:1px solid #86efac;">
                    <b>🏆 {row['Pilar']}</b><br>
                    🎯 Score: <b>{row['Score']:.1f}%</b>
                </div>
                """,
                unsafe_allow_html=True
            )

    with col2:
        st.markdown("### 🔴 Peores Pilares")
        for _, row in score_pilar.tail(3).iterrows():
            st.markdown(
                f"""
                <div style="padding:12px;border-radius:10px;background-color:#fef2f2;color:#000000;margin-bottom:8px;border:1px solid #fca5a5;">
                    <b>⚠️ {row['Pilar']}</b><br>
                    🎯 Score: <b>{row['Score']:.1f}%</b>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.divider()

    # =========================================================
    # TOP 3 MEJORES Y PEORES PREGUNTAS - COMPLETAS
    # =========================================================

    st.subheader("❓ Top 3 Mejores y Peores Preguntas")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 🟢 Mejores Preguntas")
        for _, row in score_preguntas.head(3).iterrows():
            st.markdown(
                f"""
                <div style="padding:12px;border-radius:10px;background-color:#f0fdf4;color:#000000;margin-bottom:8px;border:1px solid #86efac;">
                    <b>✅ {row['Pregunta']}</b><br>
                    🎯 Score: <b>{row['Score']:.1f}%</b>
                </div>
                """,
                unsafe_allow_html=True
            )

    with col2:
        st.markdown("### 🔴 Peores Preguntas")
        for _, row in score_preguntas.tail(3).iterrows():
            st.markdown(
                f"""
                <div style="padding:12px;border-radius:10px;background-color:#fef2f2;color:#000000;margin-bottom:8px;border:1px solid #fca5a5;">
                    <b>❌ {row['Pregunta']}</b><br>
                    🎯 Score: <b>{row['Score']:.1f}%</b>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.divider()

    # =========================================================
    # EVOLUCIÓN Y ANÁLISIS DE COACHINGS (INDEPENDIENTE DE FILTROS SUPERIORES)
    # =========================================================

    st.subheader("📈 Evolución y Comparación de Coachings")

    # USAR DATOS SIN FILTRAR PARA LOS GRÁFICOS DE ABAJO
    df_evol_base = df_respuestas.copy()

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        auditados_evol = ["Todos"] + sorted(df_evol_base["Auditado"].dropna().unique().tolist())
        auditado_evol = st.selectbox("👤 Auditado", auditados_evol, key="evol_auditado")

    with col2:
        auditores_evol = ["Todos"] + sorted(df_evol_base["Auditor"].dropna().unique().tolist())
        auditor_evol = st.selectbox("👤 Auditor", auditores_evol, key="evol_auditor")

    with col3:
        pilares_evol = ["Todos"] + sorted(df_evol_base["Pilar"].dropna().unique().tolist())
        pilar_evol = st.selectbox("🏆 Pilar", pilares_evol, key="evol_pilar")

    with col4:
        coachings_disponibles = sorted(df_evol_base["Bimestre"].dropna().unique().tolist())
        coachings_nombres = [obtener_nombre_bimestre(b) for b in coachings_disponibles]
        coaching_seleccionado = st.selectbox(
            "📅 Coaching a analizar",
            coachings_nombres,
            key="evol_coaching"
        )

    # Obtener el número de bimestre seleccionado
    bimestre_seleccionado = None
    for b in coachings_disponibles:
        if obtener_nombre_bimestre(b) == coaching_seleccionado:
            bimestre_seleccionado = b
            break

    # =========================================================
    # GRÁFICO DE EVOLUCIÓN BIMESTRAL (LÍNEA)
    # =========================================================

    # APLICAR FILTROS DE LA EVOLUCIÓN (NO LOS SUPERIORES)
    df_evol = df_evol_base.copy()

    if auditor_evol != "Todos":
            df_evol = df_evol[df_evol["Auditor"] == auditor_evol]

    if auditado_evol != "Todos":
        df_evol = df_evol[df_evol["Auditado"] == auditado_evol]

    if pilar_evol != "Todos":
        df_evol = df_evol[df_evol["Pilar"] == pilar_evol]

    # Recalcular evolución con filtros
    evolucion_filtrada = df_evol.groupby(["Bimestre", "Bimestre_Nombre"], as_index=False).agg({
        "Puntaje Final": "sum",
        "Puntos Máximo": "sum"
    })
    evolucion_filtrada["% Cumplimiento"] = (evolucion_filtrada["Puntaje Final"] / evolucion_filtrada["Puntos Máximo"] * 100).round(2)
    evolucion_filtrada = evolucion_filtrada.sort_values("Bimestre")

    if not evolucion_filtrada.empty:
        fig_line = px.line(
            evolucion_filtrada,
            x="Bimestre_Nombre",
            y="% Cumplimiento",
            markers=True,
            text=evolucion_filtrada["% Cumplimiento"].round(1).astype(str) + "%",
            title="Evolución del Score por Bimestre"
        )
        fig_line.update_traces(textposition="top center")
        fig_line.update_layout(
            xaxis_title="Bimestre",
            yaxis_title="Score (%)",
            yaxis_range=[0, 100],
            height=400
        )
        st.plotly_chart(fig_line, use_container_width=True)
    else:
        st.info("No hay datos suficientes para mostrar la evolución")

    st.divider()

    # =========================================================
    # GRÁFICO DE CASCADA - COMPARACIÓN ENTRE COACHINGS
    # =========================================================

    st.subheader("📊 Comparación entre Coachings (Cascada por Pilar)")

    if bimestre_seleccionado is not None:
        bimestre_anterior = bimestre_seleccionado - 1
        
        if bimestre_anterior in coachings_disponibles:
            
            # USAR df_evol_base (sin filtrar por los filtros superiores)
            df_coaching_actual = df_evol_base[df_evol_base["Bimestre"] == bimestre_seleccionado]
            df_coaching_anterior = df_evol_base[df_evol_base["Bimestre"] == bimestre_anterior]
            
            if not df_coaching_actual.empty and not df_coaching_anterior.empty:
                
                # APLICAR SOLO LOS FILTROS DE LA EVOLUCIÓN
                df_coaching_actual_filtrado = df_coaching_actual.copy()
                df_coaching_anterior_filtrado = df_coaching_anterior.copy()
                
                if auditado_evol != "Todos":
                    df_coaching_actual_filtrado = df_coaching_actual_filtrado[df_coaching_actual_filtrado["Auditado"] == auditado_evol]
                    df_coaching_anterior_filtrado = df_coaching_anterior_filtrado[df_coaching_anterior_filtrado["Auditado"] == auditado_evol]

                if auditor_evol != "Todos":
                                    df_coaching_actual_filtrado = df_coaching_actual_filtrado[df_coaching_actual_filtrado["Auditor"] == auditor_evol]
                                    df_coaching_anterior_filtrado = df_coaching_anterior_filtrado[df_coaching_anterior_filtrado["Auditor"] == auditor_evol]
                
                if pilar_evol != "Todos":
                    df_coaching_actual_filtrado = df_coaching_actual_filtrado[df_coaching_actual_filtrado["Pilar"] == pilar_evol]
                    df_coaching_anterior_filtrado = df_coaching_anterior_filtrado[df_coaching_anterior_filtrado["Pilar"] == pilar_evol]
                
                # Verificar que haya datos después de filtrar
                if df_coaching_actual_filtrado.empty or df_coaching_anterior_filtrado.empty:
                    st.warning(f"No hay datos para los filtros seleccionados en el Coaching {bimestre_seleccionado} o {bimestre_anterior}")
                else:
                    # Calcular score total por coaching
                    total_puntaje_actual = df_coaching_actual_filtrado["Puntaje Final"].sum()
                    total_maximo_actual = df_coaching_actual_filtrado["Puntos Máximo"].sum()
                    score_actual_total = (total_puntaje_actual / total_maximo_actual * 100) if total_maximo_actual > 0 else 0
                    
                    total_puntaje_anterior = df_coaching_anterior_filtrado["Puntaje Final"].sum()
                    total_maximo_anterior = df_coaching_anterior_filtrado["Puntos Máximo"].sum()
                    score_anterior_total = (total_puntaje_anterior / total_maximo_anterior * 100) if total_maximo_anterior > 0 else 0
                    
                    # Calcular score por Pilar
                    score_actual = df_coaching_actual_filtrado.groupby("Pilar", as_index=False).agg({
                        "Puntaje Final": "sum",
                        "Puntos Máximo": "sum"
                    })
                    score_actual["Score_Actual"] = (score_actual["Puntaje Final"] / score_actual["Puntos Máximo"] * 100).round(2)
                    
                    score_anterior = df_coaching_anterior_filtrado.groupby("Pilar", as_index=False).agg({
                        "Puntaje Final": "sum",
                        "Puntos Máximo": "sum"
                    })
                    score_anterior["Score_Anterior"] = (score_anterior["Puntaje Final"] / score_anterior["Puntos Máximo"] * 100).round(2)
                    
                    # Combinar ambos dataframes
                    df_comparacion = pd.merge(score_anterior[["Pilar", "Score_Anterior"]], 
                                             score_actual[["Pilar", "Score_Actual"]], 
                                             on="Pilar", how="outer").fillna(0)
                    
                    # Calcular diferencia por pilar
                    df_comparacion["Diferencia"] = df_comparacion["Score_Actual"] - df_comparacion["Score_Anterior"]
                    df_comparacion = df_comparacion.sort_values("Diferencia", ascending=False)
                    
                    # Construir gráfico de cascada
                    nombre_anterior = obtener_nombre_bimestre(bimestre_anterior).replace(" (Ene-Feb)", "").replace(" (Mar-Abr)", "").replace(" (May-Jun)", "").replace(" (Jul-Ago)", "").replace(" (Sep-Oct)", "").replace(" (Nov-Dic)", "")
                    nombre_actual = obtener_nombre_bimestre(bimestre_seleccionado).replace(" (Ene-Feb)", "").replace(" (Mar-Abr)", "").replace(" (May-Jun)", "").replace(" (Jul-Ago)", "").replace(" (Sep-Oct)", "").replace(" (Nov-Dic)", "")
                    
                    categorias = []
                    valores = []
                    colores = []
                    textos = []
                    bases = []
                    
                    # Barra inicial
                    categorias.append(f"{nombre_anterior}")
                    valores.append(score_anterior_total)
                    colores.append("#6c757d")
                    textos.append(f"{score_anterior_total:.1f}%")
                    bases.append(0)
                    
                    # Barras intermedias
                    acumulado = score_anterior_total
                    for _, row in df_comparacion.iterrows():
                        diferencia = row["Diferencia"]
                        pilar = row["Pilar"]
                        if diferencia >= 0:
                            categorias.append(pilar)
                            valores.append(diferencia)
                            colores.append("#22c55e")
                            textos.append(f"+{diferencia:.1f}%")
                            bases.append(acumulado)
                        else:
                            categorias.append(pilar)
                            valores.append(abs(diferencia))
                            colores.append("#ef4444")
                            textos.append(f"{diferencia:.1f}%")
                            bases.append(acumulado + diferencia)
                        acumulado += diferencia
                    
                    # Barra final
                    categorias.append(f"{nombre_actual}")
                    valores.append(score_actual_total)
                    colores.append("#3b82f6")
                    textos.append(f"{score_actual_total:.1f}%")
                    bases.append(0)
                    
                    import plotly.graph_objects as go
                    
                    fig_cascada = go.Figure()
                    for i, (cat, val, color, text, base) in enumerate(zip(categorias, valores, colores, textos, bases)):
                        fig_cascada.add_trace(go.Bar(
                            x=[cat],
                            y=[val],
                            name=cat,
                            marker_color=color,
                            text=[text],
                            textposition='outside',
                            textfont=dict(size=14, color='black'),
                            base=[base],
                            width=0.6
                        ))
                    
                    fig_cascada.update_layout(
                        height=550,
                        title=dict(text=f"Comparación: {nombre_anterior} → {nombre_actual}", font=dict(size=20)),
                        yaxis_title=dict(text="Score (%)", font=dict(size=16)),
                        xaxis_title=dict(text="", font=dict(size=16)),
                        showlegend=False,
                        xaxis=dict(tickfont=dict(size=12), tickangle=0),
                        yaxis=dict(tickfont=dict(size=14), gridcolor='#e5e7eb'),
                        bargap=0.15,
                        bargroupgap=0.05
                    )
                    
                    min_val = min([b for b in bases if b < 0]) if any(b < 0 for b in bases) else 0
                    max_val = max([b + v for b, v in zip(bases, valores)]) + 5
                    fig_cascada.update_layout(yaxis_range=[min_val - 5 if min_val < 0 else 0, max_val])
                    
                    st.plotly_chart(fig_cascada, use_container_width=True)
                    
                    with st.expander("📋 Ver detalle de la comparación"):
                        df_mostrar = df_comparacion.copy()
                        df_mostrar["Score_Anterior"] = df_mostrar["Score_Anterior"].round(1).astype(str) + "%"
                        df_mostrar["Score_Actual"] = df_mostrar["Score_Actual"].round(1).astype(str) + "%"
                        df_mostrar["Diferencia"] = df_mostrar["Diferencia"].round(1).astype(str) + "%"
                        df_mostrar.columns = ["Pilar", f"{nombre_anterior}", f"{nombre_actual}", "Cambio"]
                        
                        total_row = pd.DataFrame({
                            "Pilar": ["**TOTAL**"],
                            f"{nombre_anterior}": [f"{score_anterior_total:.1f}%"],
                            f"{nombre_actual}": [f"{score_actual_total:.1f}%"],
                            "Cambio": [f"{(score_actual_total - score_anterior_total):.1f}%"]
                        })
                        df_mostrar = pd.concat([df_mostrar, total_row], ignore_index=True)
                        st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
                
            else:
                st.warning(f"No hay datos suficientes para comparar el Coaching {bimestre_seleccionado} con el anterior")
        else:
            st.info(f"No hay un coaching anterior para comparar con {obtener_nombre_bimestre(bimestre_seleccionado)}")
    else:
        st.info("Seleccione un Coaching para ver la comparación")

# =========================================================
# CATEGORÍAS A MEJORAR
# =========================================================

elif seleccion == "🎯 Categorías a Mejorar":
    st.title("🎯 Categorías a Mejorar")

    # Cargar datos desde la hoja RESUMEN y RESPUESTAS
    datos = cargar_datos()
    df_resumen = datos["resumen"]
    df_respuestas = datos["respuestas"]

    if df_resumen.empty:
        st.warning("⚠️ No existen datos de coaching cargados")
        st.stop()

    if df_respuestas.empty:
        st.warning("⚠️ No existen respuestas de coaching cargadas")
        st.stop()

    # =========================================================
    # CONVERTIR TIPOS DE DATOS
    # =========================================================
    
    df_resumen["Puntaje Final"] = pd.to_numeric(df_resumen["Puntaje Final"], errors="coerce")
    df_resumen["Fecha"] = pd.to_datetime(df_resumen["Fecha"], errors="coerce")
    
    df_respuestas["Puntaje Final"] = pd.to_numeric(df_respuestas["Puntaje Final"], errors="coerce")
    df_respuestas["Puntos Máximo"] = pd.to_numeric(df_respuestas["Puntos Máximo"], errors="coerce")
    df_respuestas["Fecha"] = pd.to_datetime(df_respuestas["Fecha"], errors="coerce")

    # =========================================================
    # VERIFICAR NOMBRE DE LA COLUMNA DE CATEGORÍAS EN RESUMEN
    # =========================================================
    
    columna_categorias = None
    if "Categorías a Mejorar" in df_resumen.columns:
        columna_categorias = "Categorías a Mejorar"
    elif "Categoría a Mejorar" in df_resumen.columns:
        columna_categorias = "Categoría a Mejorar"
    else:
        st.error("⚠️ No se encontró la columna de categorías en la hoja RESUMEN.")
        st.stop()
    
    # =========================================================
    # NOMBRE DE LA COLUMNA DE CATEGORÍA EN RESPUESTAS
    # =========================================================
    
    columna_categoria_respuestas = "Categorías"
    
    if columna_categoria_respuestas not in df_respuestas.columns:
        st.error(f"⚠️ No se encontró la columna '{columna_categoria_respuestas}' en la hoja RESPUESTAS.")
        st.info(f"Columnas disponibles en RESPUESTAS: {list(df_respuestas.columns)}")
        st.stop()

    # =========================================================
    # DEFINIR BIMESTRES
    # =========================================================
    
    def obtener_bimestre(fecha):
        if pd.isna(fecha):
            return None
        mes = fecha.month
        if mes in [1, 2]:
            return 1
        elif mes in [3, 4]:
            return 2
        elif mes in [5, 6]:
            return 3
        elif mes in [7, 8]:
            return 4
        elif mes in [9, 10]:
            return 5
        elif mes in [11, 12]:
            return 6
        return None

    def obtener_nombre_bimestre(bimestre):
        nombres = {
            1: "Coaching 1 (Ene-Feb)",
            2: "Coaching 2 (Mar-Abr)",
            3: "Coaching 3 (May-Jun)",
            4: "Coaching 4 (Jul-Ago)",
            5: "Coaching 5 (Sep-Oct)",
            6: "Coaching 6 (Nov-Dic)"
        }
        return nombres.get(bimestre, f"Coaching {bimestre}")

    df_resumen["Bimestre"] = df_resumen["Fecha"].apply(obtener_bimestre)
    df_resumen["Bimestre_Nombre"] = df_resumen["Bimestre"].apply(obtener_nombre_bimestre)
    
    df_respuestas["Bimestre"] = df_respuestas["Fecha"].apply(obtener_bimestre)
    df_respuestas["Bimestre_Nombre"] = df_respuestas["Bimestre"].apply(obtener_nombre_bimestre)

    # =========================================================
    # FILTROS
    # =========================================================

    st.subheader("🔎 Filtros")

    col1, col2, col3 = st.columns(3)

    with col1:
        auditores_filtro = ["Todos"] + sorted(df_resumen["Auditor"].dropna().unique().tolist())
        auditor_seleccionado = st.selectbox("👤 Auditor", auditores_filtro, key="mejora_auditor")

    with col2:
        auditados_filtro = ["Todos"] + sorted(df_resumen["Auditado"].dropna().unique().tolist())
        auditado_seleccionado = st.selectbox("👤 Auditado", auditados_filtro, key="mejora_auditado")

    with col3:
        bimestres_filtro = ["Todos"] + sorted(df_resumen["Bimestre_Nombre"].dropna().unique().tolist())
        bimestre_seleccionado = st.selectbox("📅 Coaching a analizar", bimestres_filtro, key="mejora_bimestre")

    st.divider()

    # =========================================================
    # APLICAR FILTROS
    # =========================================================

    df_filtrado = df_resumen.copy()

    if auditor_seleccionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Auditor"] == auditor_seleccionado]

    if auditado_seleccionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Auditado"] == auditado_seleccionado]

    if bimestre_seleccionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Bimestre_Nombre"] == bimestre_seleccionado]

    if df_filtrado.empty:
        st.info("No hay datos con los filtros seleccionados")
        st.stop()

    # =========================================================
    # OBTENER COACHING SELECCIONADO Y ANTERIOR
    # =========================================================

    # Obtener el número de bimestre del filtro
    bimestre_actual_num = None
    for b in sorted(df_resumen["Bimestre"].dropna().unique().tolist()):
        if obtener_nombre_bimestre(b) == bimestre_seleccionado:
            bimestre_actual_num = b
            break

    if bimestre_actual_num is None:
        st.info("Seleccione un Coaching para ver la comparación")
        st.stop()

    bimestre_anterior_num = bimestre_actual_num - 1

    # Verificar que exista el coaching anterior
    if bimestre_anterior_num not in df_resumen["Bimestre"].dropna().unique().tolist():
        st.info(f"No hay un coaching anterior para comparar con {bimestre_seleccionado}")
        st.stop()

    # Obtener datos del coaching seleccionado y anterior
    df_actual = df_filtrado[df_filtrado["Bimestre"] == bimestre_actual_num]
    
    df_anterior = df_resumen[
        (df_resumen["Bimestre"] == bimestre_anterior_num)
    ]
    if auditor_seleccionado != "Todos":
        df_anterior = df_anterior[df_anterior["Auditor"] == auditor_seleccionado]
    if auditado_seleccionado != "Todos":
        df_anterior = df_anterior[df_anterior["Auditado"] == auditado_seleccionado]

    if df_actual.empty:
        st.warning(f"No hay datos para el coaching seleccionado")
        st.stop()

    if df_anterior.empty:
        st.info(f"No hay datos del coaching anterior ({obtener_nombre_bimestre(bimestre_anterior_num)}) para comparar")
        st.stop()

    # =========================================================
    # FUNCIÓN PARA CALCULAR SCORE POR CATEGORÍA DESDE RESPUESTAS
    # =========================================================
    
    def calcular_score_categoria(auditor, auditado, bimestre, categoria):
        """Calcula el score de una categoría específica para un coaching"""
        df_filtro = df_respuestas[
            (df_respuestas["Auditor"] == auditor) &
            (df_respuestas["Auditado"] == auditado) &
            (df_respuestas["Bimestre"] == bimestre) &
            (df_respuestas[columna_categoria_respuestas] == categoria)
        ]
        
        if df_filtro.empty:
            return None
        
        total_puntaje = df_filtro["Puntaje Final"].sum()
        total_maximo = df_filtro["Puntos Máximo"].sum()
        
        if total_maximo == 0:
            return 0
        
        return (total_puntaje / total_maximo * 100)

    # =========================================================
    # EXTRAER CATEGORÍAS DEL COACHING ANTERIOR
    # =========================================================
    
    def extraer_categorias(df):
        """Extrae los nombres de las categorías de la columna"""
        categorias_raw = []
        for _, row in df.iterrows():
            categorias_str = row[columna_categorias]
            if pd.isna(categorias_str) or categorias_str == "" or categorias_str == "Ninguna":
                continue
            for item in categorias_str.split(","):
                item = item.strip()
                if not item:
                    continue
                import re
                match = re.search(r'\(([\d.]+)%\)', item)
                if match:
                    nombre = item[:match.start()].strip()
                else:
                    nombre = item
                if nombre not in categorias_raw:
                    categorias_raw.append(nombre)
        return categorias_raw

    # =========================================================
    # OBTENER CATEGORÍAS DEL COACHING ANTERIOR
    # =========================================================
    
    # Obtener auditor y auditado del filtro
    auditor_filtro_val = auditor_seleccionado if auditor_seleccionado != "Todos" else None
    auditado_filtro_val = auditado_seleccionado if auditado_seleccionado != "Todos" else None
    
    if auditado_filtro_val is None:
        auditados_disponibles = df_actual["Auditado"].dropna().unique().tolist()
        if not auditados_disponibles:
            st.warning("No hay auditados disponibles")
            st.stop()
        auditado_filtro_val = auditados_disponibles[0]
    
    if auditor_filtro_val is None:
        auditores_disponibles = df_actual["Auditor"].dropna().unique().tolist()
        if not auditores_disponibles:
            st.warning("No hay auditores disponibles")
            st.stop()
        auditor_filtro_val = auditores_disponibles[0]

    # =========================================================
    # OBTENER CATEGORÍAS A MEJORAR DEL COACHING ANTERIOR
    # =========================================================
    
    categorias_anterior_list = extraer_categorias(df_anterior)
    
    if not categorias_anterior_list:
        st.success("✅ No hay categorías a mejorar en el coaching anterior")
        st.stop()

    # =========================================================
    # CALCULAR SCORES PARA CADA CATEGORÍA EN AMBOS COACHINGS
    # =========================================================
    
    categorias_anterior = {}
    categorias_actual = {}
    
    for categoria in categorias_anterior_list:
        # Score en el coaching anterior
        score_anterior = calcular_score_categoria(
            auditor_filtro_val, 
            auditado_filtro_val, 
            bimestre_anterior_num, 
            categoria
        )
        if score_anterior is not None:
            categorias_anterior[categoria] = score_anterior
        
        # Score en el coaching seleccionado
        score_actual = calcular_score_categoria(
            auditor_filtro_val, 
            auditado_filtro_val, 
            bimestre_actual_num, 
            categoria
        )
        if score_actual is not None:
            categorias_actual[categoria] = score_actual
        else:
            categorias_actual[categoria] = None

    # =========================================================
    # CREAR TABLA DE COMPARACIÓN
    # =========================================================

    st.subheader(f"📊 Comparación: {obtener_nombre_bimestre(bimestre_anterior_num)} → {obtener_nombre_bimestre(bimestre_actual_num)}")

    tabla_comparacion = []

    for categoria in categorias_anterior_list:
        score_anterior = categorias_anterior.get(categoria, 0)
        score_actual = categorias_actual.get(categoria)
        
        if score_actual is None:
            # La categoría ya no está en el coaching actual (se resolvió)
            estado = "✅ Resuelta"
            color_estado = "#22c55e"
            cambio = "N/A"
            actual_str = "N/A"
        else:
            diferencia = score_actual - score_anterior
            
            if diferencia > 0:
                if score_actual >= 75:
                    estado = "✅ Mejoró"
                    color_estado = "#22c55e"
                else:
                    estado = "🟡 En Proceso"
                    color_estado = "#eab308"
                cambio = f"+{diferencia:.1f}%"
            elif diferencia < 0:
                estado = "🔴 Empeoró"
                color_estado = "#ef4444"
                cambio = f"{diferencia:.1f}%"
            else:
                estado = "🟡 Igual"
                color_estado = "#eab308"
                cambio = "0%"
            actual_str = f"{score_actual:.1f}%"
        
        tabla_comparacion.append({
            "Categoría": categoria,
            "Anterior": f"{score_anterior:.1f}%",
            "Actual": actual_str,
            "Cambio": cambio,
            "Estado": estado,
            "color": color_estado
        })

    # Ordenar
    orden_estado = {"🔴 Empeoró": 0, "🟡 En Proceso": 1, "🟡 Igual": 2, "✅ Mejoró": 3, "✅ Resuelta": 4}
    tabla_comparacion.sort(key=lambda x: orden_estado.get(x["Estado"], 5))

    # =========================================================
    # MOSTRAR TABLA DE COMPARACIÓN
    # =========================================================

    for item in tabla_comparacion:
        st.markdown(
            f"""
            <div style="padding:12px 16px;border-radius:10px;background-color:#f8f9fa;color:#000000;margin-bottom:8px;border-left:6px solid {item['color']};">
                <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
                    <div style="font-weight:bold;font-size:15px;">📌 {item['Categoría']}</div>
                    <div style="display:flex;gap:15px;flex-wrap:wrap;align-items:center;font-size:14px;">
                        <span style="color:#6c757d;">🔙 {item['Anterior']}</span>
                        <span style="color:#6c757d;">➡️ <b>{item['Actual']}</b></span>
                        <span><b>Cambio: {item['Cambio']}</b></span>
                        <span style="background-color:{item['color']};color:white;padding:2px 12px;border-radius:12px;font-size:12px;font-weight:bold;">
                            {item['Estado']}
                        </span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    # =========================================================
    # CATEGORÍAS ACTUALES (expander)
    # =========================================================

    # =========================================================
    # CATEGORÍAS ACTUALES (expander) - CORREGIDO
    # =========================================================

    st.divider()
    
    # Obtener TODAS las categorías a mejorar del coaching actual
    categorias_actuales_completas = extraer_categorias(df_actual)
    categorias_con_score = {}
    
    for categoria in categorias_actuales_completas:
        score = calcular_score_categoria(
            auditor_filtro_val, 
            auditado_filtro_val, 
            bimestre_actual_num, 
            categoria
        )
        if score is not None:
            categorias_con_score[categoria] = score
    
    with st.expander("📋 Categorías a Mejorar del Coaching Actual", expanded=False):
        if categorias_con_score:
            # Ordenar por score de menor a mayor (las más críticas primero)
            for categoria, score in sorted(categorias_con_score.items(), key=lambda x: x[1]):
                # Verificar si es una categoría nueva (no estaba en el coaching anterior)
                es_nueva = categoria not in categorias_anterior_list
                
                if es_nueva:
                    st.markdown(
                        f"""
                        <div style="padding:8px 12px;border-radius:8px;background-color:#dbeafe;color:#000000;margin-bottom:4px;border:1px solid #3b82f6;">
                            <b>🆕 {categoria}</b> - Score: <b>{score:.1f}%</b> <span style="font-size:11px;color:#3b82f6;">(Nueva)</span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                else:
                    st.markdown(
                        f"""
                        <div style="padding:8px 12px;border-radius:8px;background-color:#fef2f2;color:#000000;margin-bottom:4px;border:1px solid #fca5a5;">
                            <b>🔴 {categoria}</b> - Score: <b>{score:.1f}%</b>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
            st.caption(f"📌 Total de categorías a mejorar en el coaching actual: {len(categorias_con_score)}")
            
            # Mostrar también las categorías que estaban en el coaching anterior pero ya no están (resueltas)
            resueltas = [c for c in categorias_anterior_list if c not in categorias_actuales_completas]
            if resueltas:
                st.caption(f"✅ Categorías resueltas (ya no están en el coaching actual): {', '.join(resueltas)}")
        else:
            st.success("✅ No hay categorías a mejorar en el coaching actual")

        # Resumen
    st.divider()
    
    total = len(tabla_comparacion)
    mejoraron = sum(1 for x in tabla_comparacion if x["Estado"] == "✅ Mejoró")
    en_proceso = sum(1 for x in tabla_comparacion if x["Estado"] == "🟡 En Proceso")
    empeoraron = sum(1 for x in tabla_comparacion if x["Estado"] == "🔴 Empeoró")
    igual = sum(1 for x in tabla_comparacion if x["Estado"] == "🟡 Igual")
    resueltas = sum(1 for x in tabla_comparacion if x["Estado"] == "✅ Resuelta")
    nuevas = sum(1 for x in tabla_comparacion if x["Estado"] == "🆕 Nueva")
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1:
        st.metric("📋 Total", total)
    with col2:
        st.metric("✅ Mejoró", mejoraron)
    with col3:
        st.metric("🟡 En Proceso", en_proceso)
    with col4:
        st.metric("🔴 Empeoró", empeoraron)
    with col5:
        st.metric("🟡 Igual", igual)
    with col6:
        st.metric("✅ Resuelta", resueltas)

# =========================================================
# HISTORIAL
# =========================================================

elif seleccion == "📂 Historial":
    st.title("📂 Historial de Coaching")

    datos = cargar_datos()
    df_respuestas = datos["respuestas"]
    df_auditados = datos["auditados"]

    if df_respuestas.empty:
        st.warning("No existen sesiones de coaching cargadas.")
        st.stop()

    # =========================================================
    # CONVERTIR TIPOS DE DATOS
    # =========================================================
    
    df_respuestas["Fecha"] = pd.to_datetime(df_respuestas["Fecha"], errors="coerce")
    df_respuestas["Puntaje Final"] = pd.to_numeric(df_respuestas["Puntaje Final"], errors="coerce")
    df_respuestas["Puntos Máximo"] = pd.to_numeric(df_respuestas["Puntos Máximo"], errors="coerce")
    df_respuestas["Peso"] = pd.to_numeric(df_respuestas["Peso"], errors="coerce")

    # =========================================================
    # AGREGAR PUESTO DEL AUDITADO
    # =========================================================
    
    if not df_auditados.empty:
        puesto_dict = dict(zip(df_auditados["Auditado"], df_auditados["Puesto"]))
        df_respuestas["Puesto"] = df_respuestas["Auditado"].map(puesto_dict)
    else:
        df_respuestas["Puesto"] = ""

    # =========================================================
    # DEFINIR BIMESTRES
    # =========================================================
    
    def obtener_bimestre(fecha):
        if pd.isna(fecha):
            return None
        mes = fecha.month
        if mes in [1, 2]:
            return 1
        elif mes in [3, 4]:
            return 2
        elif mes in [5, 6]:
            return 3
        elif mes in [7, 8]:
            return 4
        elif mes in [9, 10]:
            return 5
        elif mes in [11, 12]:
            return 6
        return None

    def obtener_nombre_bimestre(bimestre):
        nombres = {
            1: "Coaching 1 (Ene-Feb)",
            2: "Coaching 2 (Mar-Abr)",
            3: "Coaching 3 (May-Jun)",
            4: "Coaching 4 (Jul-Ago)",
            5: "Coaching 5 (Sep-Oct)",
            6: "Coaching 6 (Nov-Dic)"
        }
        return nombres.get(bimestre, f"Coaching {bimestre}")

    df_respuestas["Bimestre"] = df_respuestas["Fecha"].apply(obtener_bimestre)
    df_respuestas["Bimestre_Nombre"] = df_respuestas["Bimestre"].apply(obtener_nombre_bimestre)

    # =========================================================
    # FILTROS
    # =========================================================

    st.subheader("🔎 Filtros")

    col1, col2, col3 = st.columns(3)

    with col1:
        auditores_hist = ["Todos"] + sorted(df_respuestas["Auditor"].dropna().unique().tolist())
        auditor_hist = st.selectbox("👤 Auditor", auditores_hist, key="hist_auditor")

    with col2:
        auditados_hist = ["Todos"] + sorted(df_respuestas["Auditado"].dropna().unique().tolist())
        auditado_hist = st.selectbox("👤 Auditado", auditados_hist, key="hist_auditado")

    with col3:
        bimestres_hist = ["Todos"] + sorted(df_respuestas["Bimestre_Nombre"].dropna().unique().tolist())
        bimestre_hist = st.selectbox("📅 Coaching", bimestres_hist, key="hist_bimestre")

    # =========================================================
    # APLICAR FILTROS
    # =========================================================

    df_filtrado = df_respuestas.copy()

    if auditor_hist != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Auditor"] == auditor_hist]

    if auditado_hist != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Auditado"] == auditado_hist]

    if bimestre_hist != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Bimestre_Nombre"] == bimestre_hist]

    if df_filtrado.empty:
        st.info("No hay datos con los filtros seleccionados")
        st.stop()

    # =========================================================
    # VERIFICAR COLUMNA DE CATEGORÍA
    # =========================================================
    
    columna_categoria = None
    if "Categorías" in df_filtrado.columns:
        columna_categoria = "Categorías"
    elif "Categoría" in df_filtrado.columns:
        columna_categoria = "Categoría"
    else:
        df_filtrado["Categoría"] = ""
        columna_categoria = "Categoría"

    # =========================================================
    # PREPARAR TABLA CON LAS COLUMNAS SOLICITADAS
    # =========================================================

    tabla_historial = df_filtrado[[
        "Fecha", 
        "Auditado", 
        "Puesto", 
        "Localidad", 
        "Pregunta", 
        "Pilar", 
        columna_categoria, 
        "Puntos Máximo", 
        "Peso", 
        "Puntaje Final"
    ]].copy()
    
    if columna_categoria != "Categoría":
        tabla_historial = tabla_historial.rename(columns={columna_categoria: "Categoría"})
    
    tabla_historial["Score"] = (tabla_historial["Puntaje Final"] / tabla_historial["Puntos Máximo"] * 100).round(1)
    
    tabla_historial = tabla_historial[[
        "Fecha", 
        "Auditado", 
        "Puesto", 
        "Localidad", 
        "Pregunta", 
        "Pilar", 
        "Categoría", 
        "Puntos Máximo", 
        "Peso", 
        "Puntaje Final", 
        "Score"
    ]]
    
    tabla_historial["Fecha"] = tabla_historial["Fecha"].dt.strftime("%d-%m-%Y")
    tabla_historial = tabla_historial.sort_values("Fecha", ascending=False)

    # =========================================================
    # MOSTRAR TABLA CON SCROLL
    # =========================================================

    st.subheader("📋 Historial de Coaching")

    st.dataframe(
        tabla_historial,
        use_container_width=True,
        height=450,
        hide_index=True
    )

    # =========================================================
    # BOTÓN DE DESCARGA DIRECTA A EXCEL
    # =========================================================

    st.divider()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # Función para generar el Excel en memoria
        def generar_excel():
            buffer = io.BytesIO()
            
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                # Hoja de historial (misma tabla mostrada)
                tabla_historial.to_excel(writer, sheet_name="Historial", index=False)
                
                # Hoja de resumen por sesión (agrupado)
                resumen = df_filtrado.groupby(
                    ["Fecha", "Auditor", "Auditado", "Empresa", "Localidad", "Bimestre_Nombre"], 
                    as_index=False
                ).agg({
                    "Puntaje Final": "sum",
                    "Puntos Máximo": "sum",
                    "Pregunta": "count"
                })
                resumen["Score_Total"] = (resumen["Puntaje Final"] / resumen["Puntos Máximo"] * 100).round(1)
                resumen["Fecha"] = pd.to_datetime(resumen["Fecha"]).dt.strftime("%d-%m-%Y")
                resumen = resumen[["Fecha", "Auditor", "Auditado", "Empresa", "Localidad", "Bimestre_Nombre", "Score_Total", "Pregunta"]]
                resumen.columns = ["Fecha", "Auditor", "Auditado", "Empresa", "Localidad", "Coaching", "Score", "Preguntas"]
                resumen.to_excel(writer, sheet_name="Resumen por Sesión", index=False)
            
            buffer.seek(0)
            return buffer

        # Botón de descarga directa (sin doble clic)
        st.download_button(
            label="📥 Descargar Historial a Excel",
            data=generar_excel(),
            file_name=f"Historial_Coaching_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

# =========================================================
# MAESTROS
# =========================================================

elif seleccion == "⚙️ Maestros":
    st.title("⚙️ Gestión de Maestros")

    datos = cargar_datos()

    tablas = {
        "": "",
        "Auditores": "AUDITOR",
        "Auditados": "AUDITADO",
        "Empresas": "EMPRESAS",
        "Localidades": "LOCALIDADES",
        "Notas": "NOTAS",
        "Pilares": "PILARES",
        "Preguntas": "PREGUNTAS"
    }

    opcion_tabla = st.selectbox("Seleccionar módulo", list(tablas.keys()), index=0)

    if opcion_tabla != "":
        hoja_excel = tablas[opcion_tabla]
        df_admin = leer_hoja(hoja_excel)

        st.subheader(f"📋 {opcion_tabla}")
        st.dataframe(df_admin, use_container_width=True)

        st.divider()

        # Agregar nuevo registro
        st.subheader(f"➕ Agregar {opcion_tabla[:-1] if opcion_tabla.endswith('s') else opcion_tabla}")

        if hoja_excel == "AUDITOR":
            col1, col2 = st.columns(2)
            with col1:
                nuevo_auditor = st.text_input("Auditor")
            with col2:
                nuevo_puesto = st.text_input("Puesto")

            if st.button("➕ Agregar Auditor"):
                if nuevo_auditor and nuevo_puesto:
                    nuevo_df = pd.DataFrame({"Auditor": [nuevo_auditor], "Puesto": [nuevo_puesto]})
                    df_admin = pd.concat([df_admin, nuevo_df], ignore_index=True)
                    if guardar_hoja(df_admin, hoja_excel):
                        st.success("✅ Auditor agregado")
                        st.rerun()

        elif hoja_excel == "AUDITADO":
            col1, col2 = st.columns(2)
            with col1:
                nuevo_auditado = st.text_input("Auditado")
                nuevo_puesto = st.text_input("Puesto")
            with col2:
                lista_empresas = df_admin["Empresa"].dropna().unique().tolist() if "Empresa" in df_admin.columns else []
                nueva_empresa = st.selectbox("Empresa", [""] + lista_empresas)
                lista_auditores = df_admin["Auditor"].dropna().unique().tolist() if "Auditor" in df_admin.columns else []
                nuevo_auditor = st.selectbox("Auditor", [""] + lista_auditores)

            if st.button("➕ Agregar Auditado"):
                if nuevo_auditado and nuevo_puesto and nueva_empresa and nuevo_auditor:
                    nuevo_df = pd.DataFrame({
                        "Auditado": [nuevo_auditado],
                        "Puesto": [nuevo_puesto],
                        "Empresa": [nueva_empresa],
                        "Auditor": [nuevo_auditor]
                    })
                    df_admin = pd.concat([df_admin, nuevo_df], ignore_index=True)
                    if guardar_hoja(df_admin, hoja_excel):
                        st.success("✅ Auditado agregado")
                        st.rerun()

        elif hoja_excel == "LOCALIDADES":
            col1, col2 = st.columns(2)
            with col1:
                nueva_localidad = st.text_input("Localidad")
            with col2:
                lista_empresas = df_admin["Empresa"].dropna().unique().tolist() if "Empresa" in df_admin.columns else []
                nueva_empresa = st.selectbox("Empresa", [""] + lista_empresas)

            if st.button("➕ Agregar Localidad"):
                if nueva_localidad and nueva_empresa:
                    nuevo_df = pd.DataFrame({"Localidades": [nueva_localidad], "Empresas": [nueva_empresa]})
                    df_admin = pd.concat([df_admin, nuevo_df], ignore_index=True)
                    if guardar_hoja(df_admin, hoja_excel):
                        st.success("✅ Localidad agregada")
                        st.rerun()

        elif hoja_excel == "NOTAS":
            col1, col2 = st.columns(2)
            with col1:
                nueva_nota = st.text_input("Nueva Nota")
            with col2:
                nuevo_peso = st.number_input(
                    "Peso", 
                    min_value=0.0, 
                    max_value=1.0, 
                    step=0.01,
                    value=0.0,
                    format="%.2f"
                )

            if st.button("➕ Agregar Nota"):
                if nueva_nota:
                    if not df_admin.empty and nueva_nota in df_admin["Nota"].values:
                        st.warning(f"⚠️ La nota '{nueva_nota}' ya existe.")
                    else:
                        nuevo_df = pd.DataFrame({
                            "Nota": [nueva_nota],
                            "Peso": [nuevo_peso]
                        })
                        df_admin = pd.concat([df_admin, nuevo_df], ignore_index=True)
                        if guardar_hoja(df_admin, hoja_excel):
                            st.success(f"✅ Nota '{nueva_nota}' agregada con peso {nuevo_peso}")
                            st.rerun()
                else:
                    st.warning("⚠️ Ingrese un nombre para la nota")

        elif hoja_excel == "PILARES":
            col1, col2 = st.columns(2)
            with col1:
                nuevo_id = st.text_input("ID (ej: PROS, CIERRE, etc.)")
            with col2:
                nuevo_pilar = st.text_input("Nombre del Pilar")

            if st.button("➕ Agregar Pilar"):
                if nuevo_id and nuevo_pilar:
                    # Verificar si el ID ya existe
                    if not df_admin.empty and nuevo_id in df_admin["ID"].values:
                        st.warning(f"⚠️ El ID '{nuevo_id}' ya existe. Use otro.")
                    else:
                        nuevo_df = pd.DataFrame({
                            "ID": [nuevo_id],
                            "Pilar": [nuevo_pilar]
                        })
                        df_admin = pd.concat([df_admin, nuevo_df], ignore_index=True)
                        if guardar_hoja(df_admin, hoja_excel):
                            st.success(f"✅ Pilar '{nuevo_pilar}' agregado con ID '{nuevo_id}'")
                            st.rerun()
                else:
                    st.warning("⚠️ Complete ambos campos")

        elif hoja_excel == "PREGUNTAS":
            col1, col2 = st.columns(2)
            with col1:
                nuevo_id = st.text_input("ID")
                nueva_pregunta = st.text_area("Pregunta")
                nuevo_pilar = st.selectbox("Pilar", [""] + df_admin["Pilar"].dropna().unique().tolist() if "Pilar" in df_admin.columns else [])
            with col2:
                nuevo_puntos = st.number_input("Puntos Máximo", min_value=1, step=1)
                nueva_categoria = st.text_input("Categoría")
                nuevo_puesto = st.selectbox("Puesto", ["", "Promotor", "Merchand", "Supervisor"])

            if st.button("➕ Agregar Pregunta"):
                if nuevo_id and nueva_pregunta and nuevo_pilar and nuevo_puntos and nueva_categoria and nuevo_puesto:
                    nuevo_df = pd.DataFrame({
                        "ID": [nuevo_id],
                        "Pregunta": [nueva_pregunta],
                        "Pilar": [nuevo_pilar],
                        "Puntos Máximo": [nuevo_puntos],
                        "Categoría": [nueva_categoria],
                        "Puesto": [nuevo_puesto]
                    })
                    df_admin = pd.concat([df_admin, nuevo_df], ignore_index=True)
                    if guardar_hoja(df_admin, hoja_excel):
                        st.success("✅ Pregunta agregada")
                        st.rerun()

        else:
            # Hojas simples (una columna)
            columna = df_admin.columns[0] if not df_admin.empty else "Valor"
            nuevo_valor = st.text_input(f"Nuevo {opcion_tabla[:-1] if opcion_tabla.endswith('s') else opcion_tabla}")

            if st.button("➕ Agregar"):
                if nuevo_valor:
                    nuevo_df = pd.DataFrame({columna: [nuevo_valor]})
                    df_admin = pd.concat([df_admin, nuevo_df], ignore_index=True)
                    if guardar_hoja(df_admin, hoja_excel):
                        st.success(f"✅ {opcion_tabla[:-1]} agregado")
                        st.rerun()

        st.divider()

        # Eliminar registro - CORREGIDO
        st.subheader("🗑️ Eliminar")

        if not df_admin.empty:
            # Usar la primera columna como identificador
            columna_principal = df_admin.columns[0]
            
            # Para Pilares, mostrar ID y Pilar en el select
            if hoja_excel == "PILARES" and "ID" in df_admin.columns and "Pilar" in df_admin.columns:
                # Crear una lista con "ID - Pilar" para mostrar
                opciones = [""] + [f"{row['ID']} - {row['Pilar']}" for _, row in df_admin.iterrows()]
                eliminar_seleccion = st.selectbox("Seleccionar Pilar a eliminar", opciones, index=0)
                
                if st.button("🗑️ Eliminar"):
                    if eliminar_seleccion:
                        # Extraer el ID de la selección
                        id_a_eliminar = eliminar_seleccion.split(" - ")[0]
                        df_admin = df_admin[df_admin["ID"] != id_a_eliminar]
                        if guardar_hoja(df_admin, hoja_excel):
                            st.success(f"✅ Pilar con ID '{id_a_eliminar}' eliminado")
                            st.rerun()
                    else:
                        st.warning("⚠️ Seleccione un elemento para eliminar")
            else:
                # Para otras tablas, usar la primera columna
                lista_eliminar = [""] + df_admin[columna_principal].astype(str).tolist()
                eliminar_valor = st.selectbox(f"Seleccionar {columna_principal} a eliminar", lista_eliminar, index=0)

                if st.button("🗑️ Eliminar"):
                    if eliminar_valor:
                        df_admin = df_admin[df_admin[columna_principal].astype(str) != eliminar_valor]
                        if guardar_hoja(df_admin, hoja_excel):
                            st.success(f"✅ Valor '{eliminar_valor}' eliminado")
                            st.rerun()
                    else:
                        st.warning("⚠️ Seleccione un elemento para eliminar")
        else:
            st.info("No hay datos para eliminar")