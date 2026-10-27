"""
Aplicación Web Streamlit: Pipeline ETL Interactivo de Extracción y Carga Masiva
Transportes Díaz Acarreos Generales S.A.C.

Permite ejecutar el pipeline, visualizar el flujo paso a paso en tiempo real
y auditar interactivamente todos los datos extraídos del archivo Maestro (k.xlsx).
"""

import os
import io
import sys
import time
import zipfile
from datetime import datetime
import openpyxl
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Configuración de página
st.set_page_config(
    page_title="DAG Neumáticos - Pipeline ETL",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inclusión del directorio scripts en el path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from etl_maestro import (
    extraer_datos_maestro,
    generar_plantilla_vehiculos,
    generar_plantilla_conductores,
    generar_plantilla_asignacion_conductores,
    generar_plantilla_neumaticos,
    generar_plantilla_movimientos,
    generar_plantilla_comparativa,
    generar_plantilla_grifos
)
try:
    from generar_reporte_auditoria import crear_reporte_auditoria_excel
except ModuleNotFoundError:
    from scripts.generar_reporte_auditoria import crear_reporte_auditoria_excel

RUTA_MAESTRO_DEFECTO = os.path.join(BASE_DIR, "Archivos", "Maestro", "k.xlsx")
DIR_PLANTILLAS_DEFECTO = os.path.join(BASE_DIR, "Archivos", "Plantillas")

@st.cache_data(show_spinner=False)
def cargar_datos_maestro_cached(ruta, rango="A2:A35"):
    """Carga y cachea la extracción del archivo maestro para navegación fluida."""
    return extraer_datos_maestro(ruta, rango_resumen=rango)

# Estilos CSS personalizados
st.markdown("""
<style>
    .main-header {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border-radius: 10px;
        padding: 16px;
        border-left: 5px solid #2563EB;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .step-box {
        background: #FFFFFF;
        border-radius: 8px;
        padding: 14px 18px;
        border: 1px solid #E2E8F0;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .step-success {
        border-left: 5px solid #10B981;
    }
    /* Ocultar botón Deploy de la esquina superior derecha */
    .stAppDeployButton, .stDeployButton, [data-testid="stDeployButton"], [data-testid="stToolbarActions"] {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* Ocultar menú hamburguesa y barra superior */
    #MainMenu {
        visibility: hidden !important;
        display: none !important;
    }
    header [data-testid="stToolbar"] {
        visibility: hidden !important;
        display: none !important;
    }

    /* Ocultar pie de página 'Made with Streamlit' */
    footer, [data-testid="stFooter"] {
        visibility: hidden !important;
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)


# --- Barra Lateral (Configuración) ---
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/truck.png", width=70)
    st.title("Configuración")
    st.markdown("**DAG Neumáticos**  \n*Transportes Díaz Acarreos Generales*")
    st.divider()

    st.subheader("Rutas de Archivos")
    ruta_maestro = st.text_input("Archivo Maestro (k.xlsx):", value=RUTA_MAESTRO_DEFECTO)
    dir_plantillas = st.text_input("Carpeta Plantillas Destino:", value=DIR_PLANTILLAS_DEFECTO)

    st.subheader("Filtro de Carga")
    rango_filas = st.text_input("Rango en hoja 'Resumen':", value="A2:A35", help="Celdas de las unidades activas en la pestaña Resumen.")

    st.divider()

    # Estado de los archivos en disco
    maestro_ok = os.path.exists(ruta_maestro)
    plantillas_ok = os.path.exists(dir_plantillas)

    if maestro_ok:
        size_mb = os.path.getsize(ruta_maestro) / (1024 * 1024)
        mtime = datetime.fromtimestamp(os.path.getmtime(ruta_maestro)).strftime('%d/%m/%Y %H:%M')
        st.success(f"Maestro detectado ({size_mb:.1f} MB)\nÚlt. modif: {mtime}")
    else:
        st.error("No se encuentra el archivo maestro")

    if plantillas_ok:
        archivos_p = [f for f in os.listdir(dir_plantillas) if f.endswith('.xlsx') and not f.startswith('~$')]
        if "Reporte_Auditoria_Conciliacion_Neumaticos_DAG.xlsx" in archivos_p:
            st.success(f"Archivos en destino: {len(archivos_p)} (7 oficiales + Reporte Auditoría)")
        else:
            st.info(f"Plantillas en destino: {len(archivos_p)} archivos")
    else:
        st.warning("Directorio de plantillas no existe")

    st.divider()
    btn_ejecutar = st.button("🚀 Ejecutar Pipeline ETL", type="primary", width="stretch")


# --- Encabezado Principal ---
col_logo, col_titulo = st.columns([1, 10])
with col_logo:
    st.write(" ")
    st.write(" ")
    st.markdown("<h1>🚚</h1>", unsafe_allow_html=True)
with col_titulo:
    st.markdown("<div class='main-header'>Pipeline ETL - Extracción de Maestro a Plantillas</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-header'>Transportes Díaz Acarreos Generales S.A.C. | Flota Operativa Activa (Resumen A2:A35)</div>", unsafe_allow_html=True)


# Función para cargar los datos en sesión
def ejecutar_flujo_etl(ruta_m, dir_p, rango_r):
    """Ejecuta el flujo ETL con retroalimentación visual paso a paso y control de errores."""
    contenedor_flujo = st.container()
    
    with contenedor_flujo:
        st.subheader("🔄 Flujo de Ejecución del Pipeline")
        progreso = st.progress(0)
        status_box = st.empty()

        try:
            # Paso 1: Verificación
            status_box.info("🔍 Paso 1/6: Verificando integridad de archivos y entorno...")
            progreso.progress(15)
            time.sleep(0.2)
            if not os.path.exists(ruta_m):
                status_box.error(f"Error: No existe el archivo {ruta_m}")
                return None
            
            # Paso 2: Extracción del Maestro
            status_box.info(f"📑 Paso 2/6: Leyendo pestaña 'Resumen' ({rango_r}) y extrayendo datos operativos...")
            progreso.progress(35)
            datos = extraer_datos_maestro(ruta_m, rango_resumen=rango_r)

            # Paso 3: Vehículos y Conductores
            status_box.info("🚛 Paso 3/6: Generando plantillas de Vehículos y Conductores...")
            progreso.progress(55)
            generar_plantilla_vehiculos(dir_p, datos['tractos'], datos['cisternas'])
            generar_plantilla_conductores(dir_p, datos['conductores'])
            generar_plantilla_asignacion_conductores(dir_p, datos['asignaciones_conductores'], datos['tractos'], datos['conductores'])

            # Paso 4: Neumáticos y Movimientos
            status_box.info("🛞 Paso 4/7: Generando plantillas de Neumáticos e inventario de montajes...")
            progreso.progress(70)
            generar_plantilla_neumaticos(dir_p, datos['neumaticos'])
            generar_plantilla_movimientos(dir_p, datos['movimientos'])

            # Paso 5: Combustible y Grifos
            status_box.info("⛽ Paso 5/7: Generando tabla comparativa de combustible y grifos...")
            progreso.progress(85)
            generar_plantilla_comparativa(dir_p, datos['tractos'])
            generar_plantilla_grifos(dir_p)

            # Paso 6: Reporte de Auditoría y Conciliación
            status_box.info("📊 Paso 6/7: Generando Reporte de Auditoría y Conciliación de Neumáticos...")
            progreso.progress(95)
            reporte_path = os.path.join(dir_p, 'Reporte_Auditoria_Conciliacion_Neumaticos_DAG.xlsx')
            try:
                crear_reporte_auditoria_excel(datos, reporte_path)
                dir_rep_alt = os.path.join(BASE_DIR, "Archivos", "Reportes")
                if os.path.exists(dir_rep_alt):
                    crear_reporte_auditoria_excel(datos, os.path.join(dir_rep_alt, 'Reporte_Auditoria_Conciliacion_Neumaticos_DAG.xlsx'))
            except Exception as e_rep:
                st.warning(f"Aviso al generar reporte de auditoría: {e_rep}")

            # Paso 7: Finalización
            progreso.progress(100)
            status_box.success("✅ ¡Pipeline ejecutado exitosamente! Las 7 plantillas oficiales y el Reporte de Auditoría han sido generados y actualizados.")
            cargar_datos_maestro_cached.clear()
            time.sleep(0.3)
            return datos
        except Exception as e:
            status_box.error(f"❌ Error durante la ejecución del pipeline: {e}")
            st.error(f"Detalle técnico del error: {e}")
            return None


# Si el usuario presiona el botón o ya tenemos datos en memoria
if btn_ejecutar:
    with st.spinner("Procesando datos del maestro..."):
        datos_cargados = ejecutar_flujo_etl(ruta_maestro, dir_plantillas, rango_filas)
        if datos_cargados:
            st.session_state['datos_etl'] = datos_cargados
            st.session_state['ultima_ejecucion'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

# Cargar automáticamente los datos existentes si no se ha ejecutado el botón
if 'datos_etl' not in st.session_state:
    if maestro_ok:
        with st.spinner("Cargando vista previa de datos extraídos..."):
            st.session_state['datos_etl'] = cargar_datos_maestro_cached(ruta_maestro, rango_filas)
            st.session_state['ultima_ejecucion'] = "Datos en memoria"


# --- Visualización de Métricas y Detalles ---
if 'datos_etl' in st.session_state and st.session_state['datos_etl']:
    datos = st.session_state['datos_etl']

    tractos_df = pd.DataFrame(list(datos['tractos'].values()))
    cisternas_df = pd.DataFrame(list(datos['cisternas'].values()))
    vehiculos_df = pd.concat([tractos_df, cisternas_df], ignore_index=True)
    
    conductores_df = pd.DataFrame(list(datos['conductores'].values()))
    asignaciones_df = pd.DataFrame(list(datos['asignaciones_conductores'].values()))
    neumaticos_df = pd.DataFrame(list(datos['neumaticos'].values()))
    movimientos_df = pd.DataFrame(datos['movimientos'])

    # Asegurar tipos homogéneos en códigos de neumáticos para evitar advertencias de serialización
    if not neumaticos_df.empty:
        neumaticos_df['codigo'] = neumaticos_df['codigo'].astype(str)
    if not movimientos_df.empty:
        movimientos_df['codigo'] = movimientos_df['codigo'].astype(str)

    # Métricas Principales (KPIs)
    st.markdown("### 📊 Métricas de la Flota Operativa")
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("🚛 Tractocamiones", len(tractos_df), delta="Activos")
    kpi2.metric("🚚 Semirremolques", len(cisternas_df), delta="Cisternas")
    kpi3.metric("👨‍✈️ Conductores", len(conductores_df), delta="Validados")
    kpi4.metric("🛞 Neumáticos", len(neumaticos_df), delta="Únicos")
    kpi5.metric("🔄 Movimientos", len(movimientos_df), delta="Montajes")
    kpi6.metric("⛽ Combustible", len(tractos_df) * 3, delta="Registros")

    st.divider()

    # Pestañas Interactivas
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🚛 Flota y Vehículos",
        "👨‍✈️ Conductores",
        "🛞 Neumáticos",
        "🔄 Esquema de Montaje",
        "⛽ Combustible y Grifos",
        "📁 Descarga de Plantillas"
    ])

    # --- TAB 1: VEHÍCULOS ---
    with tab1:
        st.subheader("Catálogo de Vehículos Extraídos")
        c_filt1, c_filt2, c_filt3 = st.columns([2, 2, 6])
        with c_filt1:
            filtro_tipo = st.multiselect("Filtrar Tipo:", options=vehiculos_df['tipo'].unique(), default=vehiculos_df['tipo'].unique())
        with c_filt2:
            filtro_marca = st.multiselect("Filtrar Marca:", options=vehiculos_df['marca'].unique(), default=vehiculos_df['marca'].unique())

        df_veh_filtrado = vehiculos_df[
            (vehiculos_df['tipo'].isin(filtro_tipo)) &
            (vehiculos_df['marca'].isin(filtro_marca))
        ]

        cols_mostrar = ['placa', 'tipo', 'marca', 'modelo', 'tipo_aro', 'estado', 'km', 'carroceria']
        col_cfg = {
            "placa": "Placa",
            "tipo": "Tipo de Unidad",
            "marca": "Marca",
            "modelo": "Modelo",
            "tipo_aro": "Aro",
            "estado": "Estado",
            "km": st.column_config.NumberColumn("Kilometraje", format="%.1f km"),
            "carroceria": "Carrocería"
        }
        if 'tanque_lateral_izquierdo' in df_veh_filtrado.columns:
            cols_mostrar.extend(['tanque_lateral_izquierdo', 'tanque_lateral_derecho', 'tanque_mochila_auxiliar'])
            col_cfg["tanque_lateral_izquierdo"] = st.column_config.NumberColumn("Tanque Izq (gal)", format="%d gal")
            col_cfg["tanque_lateral_derecho"] = st.column_config.NumberColumn("Tanque Der (gal)", format="%d gal")
            col_cfg["tanque_mochila_auxiliar"] = st.column_config.NumberColumn("Mochila (gal)", format="%d gal")

        st.dataframe(
            df_veh_filtrado[cols_mostrar],
            column_config=col_cfg
        )

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            fig_marcas = px.pie(tractos_df, names='marca', title='Distribución de Tractocamiones por Marca', hole=0.4, color_discrete_sequence=px.colors.qualitative.Safe)
            st.plotly_chart(fig_marcas)
        with col_g2:
            fig_km = px.bar(tractos_df.sort_values('km', ascending=False).head(15), x='placa', y='km', title='Top 15 Unidades con Mayor Kilometraje', color='marca', text_auto='.2s')
            st.plotly_chart(fig_km)

    # --- TAB 2: CONDUCTORES ---
    with tab2:
        st.subheader("Conductores y Asignaciones Activas")
        c_c1, c_c2 = st.columns([1, 1])

        with c_c1:
            st.markdown("**Lista de Conductores Validados**")
            st.dataframe(conductores_df, column_config={
                "dni": "DNI",
                "nombre": "Nombre del Conductor",
                "estado": "Estado"
            })

        with c_c2:
            st.markdown("**Asignación de Conductor a Tractocamión**")
            st.dataframe(asignaciones_df, column_config={
                "placa": "Tractocamión",
                "dni": "DNI",
                "nombre": "Conductor Asignado"
            })

    # --- TAB 3: NEUMÁTICOS ---
    with tab3:
        n_neu_tot = len(neumaticos_df)
        n_mov_tot = len(movimientos_df)
        n_dup_tot = n_mov_tot - n_neu_tot
        st.subheader(f"Catálogo Maestro de Neumáticos ({n_neu_tot} Activos Únicos - Opción B)")
        with st.expander(f"📋 Ver Conciliación Matemática y Auditoría (816 Teóricos vs {n_mov_tot} Montajes vs {n_neu_tot} Únicos)", expanded=False):
            st.markdown(f"""
            **Conciliación Matemática de la Flota Activa:**
            * **Capacidad Teórica (34 unidades x 24 posiciones):** **816 cupos** (11 en tracto + 13 en cisterna).
            * **Montajes Físicos Extraídos:** **{n_mov_tot} posiciones** reales instaladas.
            * **Posiciones Faltantes / Sin Código (12 faltantes en 7 unidades):**
              * `ACF-943`: Faltan 2 (Tracto RT y Cisterna RS vacíos).
              * `AKO-813`: Falta 1 (Tracto RT anotado como 'Sin código').
              * `AMY-787`: Falta 1 (Tracto RT vacío).
              * `AUW-702`: Falta 1 (Cisterna RS anotada como 'SIN CÓDIGO').
              * `T5Q-937`: Falta 1 (Cisterna RS vacía).
              * `T5R-834`: Falta 1 (Cisterna RS vacía).
              * `T8U-906`: Faltan 5 (Tracto P3, P7, P8 y RT con 'Sin codigo'; Cisterna RS vacía).
            * **Neumáticos Únicos ({n_neu_tot} activos - Opción B):** En `k.xlsx` existen **{n_dup_tot} códigos seriales repetidos** ({n_dup_tot * 2} montajes). En la base de datos relacional de DAG Neumáticos, el catálogo maestro consolida los {n_neu_tot} códigos únicos para preservar la llave primaria (*Primary Key*), mientras que la plantilla de movimientos mantiene los {n_mov_tot} montajes físicos.
            """)
        n_col1, n_col2, n_col3 = st.columns(3)
        with n_col1:
            filtro_t_tipo = st.multiselect("Aplicación:", options=neumaticos_df['tipo'].unique(), default=neumaticos_df['tipo'].unique())
        with n_col2:
            filtro_t_marca = st.multiselect("Marca:", options=neumaticos_df['marca'].unique(), default=neumaticos_df['marca'].unique()[:5])
        with n_col3:
            filtro_t_medida = st.multiselect("Medida:", options=neumaticos_df['medida'].unique(), default=neumaticos_df['medida'].unique())

        df_neu_filtrado = neumaticos_df[
            (neumaticos_df['tipo'].isin(filtro_t_tipo)) &
            (neumaticos_df['marca'].isin(filtro_t_marca)) &
            (neumaticos_df['medida'].isin(filtro_t_medida))
        ]

        st.dataframe(
            df_neu_filtrado[['codigo', 'marca', 'medida', 'tipo', 'diseno', 'condicion', 'prof_int', 'prof_cent', 'prof_ext', 'psi']],
            column_config={
                "codigo": "Código",
                "marca": "Marca",
                "medida": "Medida",
                "tipo": "Tipo de Neumático",
                "diseno": "Diseño Banda",
                "condicion": "Condición",
                "prof_int": st.column_config.NumberColumn("Prof. Int (mm)", format="%.1f"),
                "prof_cent": st.column_config.NumberColumn("Prof. Cent (mm)", format="%.1f"),
                "prof_ext": st.column_config.NumberColumn("Prof. Ext (mm)", format="%.1f"),
                "psi": "PSI"
            }
        )

        col_ng1, col_ng2 = st.columns(2)
        with col_ng1:
            fig_tipo_neu = px.bar(neumaticos_df['tipo'].value_counts().reset_index(), x='tipo', y='count', title='Neumáticos por Tipo de Aplicación', color='tipo', text='count')
            st.plotly_chart(fig_tipo_neu)
        with col_ng2:
            fig_marcas_neu = px.pie(neumaticos_df, names='marca', title='Participación de Marcas de Neumáticos', hole=0.3)
            st.plotly_chart(fig_marcas_neu)

    # --- TAB 4: ESQUEMA DE MONTAJE ---
    with tab4:
        st.subheader("Auditoría de Montaje de Neumáticos por Unidad")
        st.markdown("Seleccione una unidad para visualizar la distribución exacta de neumáticos en cada una de sus posiciones:")

        unidad_seleccionada = st.selectbox("Seleccione Tractocamión:", options=sorted(tractos_df['placa'].unique()))
        tracto_info = datos['tractos'].get(unidad_seleccionada, {})
        cisterna_asociada = tracto_info.get('cisterna')
        fecha_insp = tracto_info.get('fecha_inspeccion')
        km_val = tracto_info.get('km', 0.0)

        info_recuadro = f"🔍 **Recuadro más reciente verificado:** Odómetro de Inspección: **{km_val:,.1f} km**"
        if fecha_insp:
            info_recuadro += f" | Fecha de Inspección: **{fecha_insp}**"
        st.info(info_recuadro)

        neu_detalles = neumaticos_df[['codigo', 'marca', 'medida', 'tipo', 'diseno', 'prof_int', 'prof_ext', 'psi']].drop_duplicates('codigo')

        col_esq1, col_esq2 = st.columns(2)

        with col_esq1:
            st.markdown(f"#### 🚛 Tractocamión: `{unidad_seleccionada}`")
            movs_tracto = movimientos_df[movimientos_df['placa'] == unidad_seleccionada]
            if not movs_tracto.empty:
                movs_tracto_det = pd.merge(movs_tracto, neu_detalles, on='codigo', how='left')
                st.dataframe(
                    movs_tracto_det[['posicion', 'codigo', 'marca', 'medida', 'tipo', 'diseno', 'prof_int', 'prof_ext']],
                    column_config={
                        "posicion": "Posición",
                        "codigo": "Código",
                        "marca": "Marca",
                        "medida": "Medida",
                        "tipo": "Aplicación",
                        "diseno": "Diseño",
                        "prof_int": st.column_config.NumberColumn("Prof. Int", format="%.1f mm"),
                        "prof_ext": st.column_config.NumberColumn("Prof. Ext", format="%.1f mm")
                    }
                )
            else:
                st.info("Sin registros de movimientos.")

        with col_esq2:
            st.markdown(f"#### 🚚 Semirremolque (Cisterna): `{cisterna_asociada or 'Sin Cisterna'}`")
            if cisterna_asociada:
                movs_cisterna = movimientos_df[movimientos_df['placa'] == cisterna_asociada]
                if not movs_cisterna.empty:
                    movs_cisterna_det = pd.merge(movs_cisterna, neu_detalles, on='codigo', how='left')
                    st.dataframe(
                        movs_cisterna_det[['posicion', 'codigo', 'marca', 'medida', 'tipo', 'diseno', 'prof_int', 'prof_ext']],
                        column_config={
                            "posicion": "Posición",
                            "codigo": "Código",
                            "marca": "Marca",
                            "medida": "Medida",
                            "tipo": "Aplicación",
                            "diseno": "Diseño",
                            "prof_int": st.column_config.NumberColumn("Prof. Int", format="%.1f mm"),
                            "prof_ext": st.column_config.NumberColumn("Prof. Ext", format="%.1f mm")
                        }
                    )
                else:
                    st.info("Sin registros de cisterna.")
            else:
                st.info("Esta unidad no tiene cisterna asignada en la pestaña 'Resumen'.")

    # --- TAB 5: COMBUSTIBLE Y GRIFOS ---
    with tab5:
        st.subheader("Tabla Comparativa de Combustible y Grifos")
        c_comb1, c_comb2 = st.columns([3, 2])

        with c_comb1:
            st.markdown("**Tabla Comparativa por Vehículo**")
            filas_comp = []
            for p in sorted(tractos_df['placa'].unique()):
                m = datos['tractos'][p]['marca']
                g = 223 if m == 'Volvo' else 200
                filas_comp.append({'PLACA': p, 'PRODUCTO': 'GASOLINA REGULAR', 'GALONES': None, 'ESTADO': 'ACTIVO'})
                filas_comp.append({'PLACA': p, 'PRODUCTO': 'GASOLINA PREMIUM', 'GALONES': None, 'ESTADO': 'ACTIVO'})
                filas_comp.append({'PLACA': p, 'PRODUCTO': 'DIESEL B5 S-50', 'GALONES': g, 'ESTADO': 'ACTIVO'})
            st.dataframe(pd.DataFrame(filas_comp))

        with c_comb2:
            st.markdown("**Estaciones de Despacho (Grifos Autorizados)**")
            grifos_df = pd.DataFrame([
                ['GRIFO PRIMAX PIURA CENTRO', 'PRIMAX', 'PIURA', 'AV. SANCHEZ CERRO 1234', 'ACTIVO'],
                ['ESTACION REPSOL TALARA SUR', 'REPSOL', 'TALARA', 'PANAMERICANA NORTE KM 1080', 'ACTIVO'],
                ['PATIO CENTRAL DAG PAITA', 'PATIO / PROPIO', 'PAITA', 'ZONA INDUSTRIAL MZ. C LOTE 4', 'ACTIVO'],
                ['GRIFO PETROPERU CHICLAYO', 'PETROPERÚ', 'CHICLAYO', 'AV. VICTOR RAUL HAYA DE LA TORRE 500', 'ACTIVO'],
                ['GRIFO SHELL LIMA NORTE', 'SHELL', 'LIMA', 'PANAMERICANA NORTE KM 22', 'ACTIVO']
            ], columns=['NOMBRE_GRIFO', 'PROVEEDOR', 'CIUDAD', 'DIRECCION', 'ESTADO'])
            st.dataframe(grifos_df)

    # --- TAB 6: DESCARGA DE ARCHIVOS ---
    with tab6:
        st.subheader("📁 Descarga Directa de Plantillas Oficiales")
        st.markdown("Los siguientes archivos fueron generados y están listos para ser importados en la aplicación DAG Neumáticos:")

        plantillas_info = [
            ("Plantilla_Registro_Masivo_Vehiculos_DAG.xlsx", "Plantilla de Vehículos", "68 unidades (34 tractos + 34 cisternas)"),
            ("Plantilla_Registro_Masivo_Conductores_DAG.xlsx", "Plantilla de Conductores", "30 conductores validados"),
            ("Plantilla_Registro_Masivo_Asignacion_Conductores_DAG.xlsx", "Asignación de Conductores", "34 asignaciones activas"),
            ("Plantilla_Registro_Masivo_Neumaticos_DAG.xlsx", "Plantilla de Neumáticos", "780 neumáticos únicos"),
            ("Plantilla_Asignacion_Movimiento_Neumaticos_DAG.xlsx", "Asignación y Montajes", "804 montajes en posiciones P1..P22/RT/RS"),
            ("Plantilla_Registro_Masivo_Comparativa_DAG.xlsx", "Tabla Comparativa Combustible", "102 registros"),
            ("Plantilla_Registro_Masivo_Grifos_DAG.xlsx", "Catálogo de Grifos", "5 estaciones de despacho"),
            ("Reporte_Auditoria_Conciliacion_Neumaticos_DAG.xlsx", "📊 Reporte de Auditoría y Conciliación", "Informe de 4 hojas: Justificación Opción B, 12 Faltantes, 24 Duplicados y Estado de las 34 Unidades")
        ]

        # Botón para descargar todas las plantillas en un archivo ZIP
        zip_buffer = io.BytesIO()
        archivos_encontrados = 0
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for fname, _, _ in plantillas_info:
                fpath = os.path.join(dir_plantillas, fname)
                if os.path.exists(fpath):
                    zip_file.write(fpath, arcname=fname)
                    archivos_encontrados += 1

        if archivos_encontrados > 0:
            st.download_button(
                label="📦 Descargar Todas las Plantillas en ZIP (.zip)",
                data=zip_buffer.getvalue(),
                file_name="Plantillas_DAG_Neumaticos_Oficiales.zip",
                mime="application/zip",
                type="primary"
            )
            st.divider()

        for fname, title, desc in plantillas_info:
            fpath = os.path.join(dir_plantillas, fname)
            if os.path.exists(fpath):
                col_d1, col_d2 = st.columns([4, 1])
                with col_d1:
                    st.markdown(f"**{title}** (`{fname}`)  \n*{desc}*")
                with col_d2:
                    with open(fpath, "rb") as f:
                        st.download_button(
                            label="⬇️ Descargar",
                            data=f.read(),
                            file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{fname}"
                        )
                st.divider()

else:
    st.warning("No se han cargado datos. Verifique la ruta del archivo maestro y presione 'Ejecutar Pipeline ETL'.")
