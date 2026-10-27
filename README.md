# Pipeline ETL con Docker, Streamlit y Python - DAG Neumáticos

**Transportes Díaz Acarreos Generales S.A.C.**

Reingeniería del pipeline de datos utilizando **Docker + Streamlit + Python**, eliminando la sobrecarga de Apache Airflow y proporcionando un contenedor ligero, rápido e interactivo que extrae los datos del archivo maestro (`k.xlsx`), muestra el flujo paso a paso en tiempo real y permite auditar y descargar las 7 plantillas oficiales de importación masiva.

---

## 📁 Estructura del Proyecto

```text
Pipeline_app-web-dag-neumaticos/
├── .github/
│   └── workflows/
│       ├── release.yml                                       # Generador de Releases SemVer automáticos
│       └── ci.yml                                            # Verificación CI de sintaxis y ejecución ETL
├── archives/                                                 # Carpeta de respaldos y snapshots generados
├── Archivos/
│   ├── Maestro/
│   │   └── k.xlsx                                            # Archivo Excel maestro origen
│   ├── Plantillas/                                           # 7 plantillas oficiales destino
│   │   ├── Plantilla_Registro_Masivo_Vehiculos_DAG.xlsx
│   │   ├── Plantilla_Registro_Masivo_Conductores_DAG.xlsx
│   │   ├── Plantilla_Registro_Masivo_Asignacion_Conductores_DAG.xlsx
│   │   ├── Plantilla_Registro_Masivo_Neumaticos_DAG.xlsx
│   │   ├── Plantilla_Asignacion_Movimiento_Neumaticos_DAG.xlsx
│   │   ├── Plantilla_Registro_Masivo_Comparativa_DAG.xlsx
│   │   └── Plantilla_Registro_Masivo_Grifos_DAG.xlsx
│   ├── Plantillas_Backup/                                    # Copia de seguridad original
│   └── Reportes/                                             # Reportes de auditoría y conciliación
├── scripts/
│   ├── etl_maestro.py                                        # Motor ETL con filtro Resumen!A2:A35
│   ├── generar_reporte_auditoria.py                          # Motor de Auditoría y Conciliación Opción B
│   ├── auto_push/                                            # Automatización de Git, Gemini y Respaldos
│   │   ├── auto_push.ps1 / .bat / .py                        # Commits convencionales con IA y Push
│   │   ├── archify.ps1 / .bat                                # Empaquetado de plantillas y snapshots
│   │   ├── api_keys.txt                                      # Configuración de Gemini API
│   │   └── README.md
│   ├── diagramas/                                            # Arquitectura interactiva y diagramas
│   │   ├── arquitectura.json / .html                         # Diagrama de arquitectura interactivo (Archify)
│   │   ├── flujo_etl_neumaticos.html                         # Flujo interactivo del pipeline ETL
│   │   ├── secuencia_etl.mmd                                 # Secuencia Mermaid
│   │   ├── render_arquitectura.ps1 / .bat                    # Compilador con Archify CLI
│   │   └── README.md
│   └── __init__.py
├── app.py                                                    # Aplicación Web Streamlit interactiva
├── Dockerfile                                                # Contenedor Streamlit + Python 3.11
├── docker-compose.yaml                                       # Orquestador Docker Compose
├── iniciar_docker_streamlit.bat                              # Iniciar en Docker en 1 clic
├── detener_docker_streamlit.bat                              # Detener contenedor en 1 clic
├── ver_logs_docker.bat                                       # Ver logs del contenedor
├── iniciar_streamlit.bat                                     # Opción de arranque local (sin Docker)
├── ejecutar_etl_local.bat                                    # Opción de extracción por consola
└── requirements.txt                                          # Dependencias Python
```

---

## 🚀 Inicio Rápido con Docker (1 Clic)

1. Haga doble clic en el archivo **`iniciar_docker_streamlit.bat`**  
   *(o ejecute `docker compose up -d --build` en la consola)*.
2. La aplicación web se abrirá automáticamente en:  
   👉 **[http://localhost:8502](http://localhost:8502)**

Para detener el contenedor en cualquier momento:
- Haga doble clic en **`detener_docker_streamlit.bat`** *(o `docker compose down`)*.

---

## 🔄 Flujo del Pipeline en la Aplicación Web

La aplicación ejecuta y visualiza el pipeline etapa por etapa:

```text
[1] 🔍 Verificación de Integridad de Archivos (Maestro y Plantillas)
       │
       ▼
[2] 📑 Lectura de Flota Activa desde Resumen!A2:A35 (34 Unidades)
       │
       ▼
[3] 🚛 Extracción de Tractocamiones y Cisternas (68 Vehículos)
       │
       ▼
[4] 👨‍✈️ Extracción de Conductores con DNI y Asignaciones (30 Conductores, 34 Asignaciones)
       │
       ▼
[5] 🛞 Extracción y Clasificación de Neumáticos (780 Neumáticos Únicos)
       │
       ▼
[6] 🔄 Generación de Montajes Físicos por Posición (804 Movimientos P1..P22/RT/RS)
       │
       ▼
[7] ⛽ Generación de Tabla Comparativa de Combustible (102 Registros) y Grifos (5)
       │
       ▼
[8] 💾 Actualización de las 7 Plantillas Oficiales y Generación del Reporte de Auditoría
```

---

## 📊 Vistas y Herramientas del Dashboard

- **Panel Superior:** Botón `🚀 Ejecutar Pipeline ETL`, selector de rango de unidades de la pestaña 'Resumen', barra de progreso dinámica y 6 tarjetas de métricas en tiempo real (KPIs).
- **Pestaña 1 (🚛 Vehículos):** Tabla con buscador por placa, filtros por tipo y marca, gráfico de composición de flota y ranking de kilometrajes.
- **Pestaña 2 (👨‍✈️ Conductores):** Lista con DNI de 8 dígitos y relación de asignación a cada tractocamión.
- **Pestaña 3 (🛞 Neumáticos):** Inventario limpio con filtros por tipo (`DIRECCION`, `TRACCION`, `ARRASTRE`), marca y medida, con gráficos interactivos y banner desplegable de conciliación matemática (816 Teóricos vs 804 Montajes vs 780 Únicos).
- **Pestaña 4 (🔄 Esquema de Montaje por Unidad):** Selector de tractocamión que renderiza la distribución detallada de neumáticos montados en el tracto (`P1`..`P10`/`RT`) y en su cisterna (`P11`..`P22`/`RS`), con marca, medida, tipo, diseño, profundidad de cocada, odómetro y fecha de última inspección.
- **Pestaña 5 (⛽ Combustible y Grifos):** Tabla comparativa de combustible y catálogo de 5 estaciones de despacho autorizadas.
- **Pestaña 6 (📁 Descarga de Plantillas):** Botón para descargar el paquete consolidado en archivo comprimido **ZIP (`Plantillas_DAG_Neumaticos_Oficiales.zip`)** o descarga individual de cada una de las 7 plantillas Excel generadas, más el **`Reporte_Auditoria_Conciliacion_Neumaticos_DAG.xlsx`** (4 hojas detalladas: Justificación Opción B, 12 Faltantes, 24 Duplicados y Estado de las 34 Unidades).
