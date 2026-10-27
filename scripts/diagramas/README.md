# 📐 Diagramas de Arquitectura y Flujos del Pipeline ETL

Este directorio centraliza los diagramas de arquitectura interactiva, modelos de flujo de datos y secuencias de interacción para el proyecto **Pipeline ETL DAG Neumáticos (Streamlit + Docker)**.

---

## 📂 Recursos Disponibles

### 1. Diagrama de Arquitectura Interactivo con Archify
- **`arquitectura.html`**:
  - Renderizado autónomo interactivo compilado con **Archify CLI**.
  - Soporte de modo oscuro/claro, zoom continuo, pan interactivo y exportación (PNG/SVG).
  - Incluye **13 componentes** organizados en 4 zonas de seguridad y arquitectura:
    1. *Zona de Operador y Cliente:* Navegador Web (:8502), Lanzadores Batch (1-Clic) y Auto-Push Git.
    2. *Zona de Aplicación y Contenedores (Docker):* Servidor Streamlit/Uvicorn, Motor ETL Maestro (8 etapas), Motor de Auditoría y Archify Packager.
    3. *Zona de Datos y Plantillas Excel:* Archivo Maestro (`k.xlsx`), 7 Plantillas Oficiales, Reporte de Auditoría y Depósito de Respaldos `archives/`.
    4. *Zona Externa GitHub & CI/CD:* Repositorio GitHub y GitHub Actions (SemVer Release).
  - **4 fichas técnicas detalladas** accesibles en la parte inferior del lienzo.
- **`arquitectura.json`**:
  - Especificación tipada estandarizada compatible con el esquema Archify v1.

### 2. Flujo del Pipeline y Secuencias (Mermaid)
- **`flujo_etl_neumaticos.html`**:
  - Visualizador interactivo completo con soporte de tema claro/oscuro.
  - Describe las 4 fases maestras: Extracción Estricta, Vehículos con 11 columnas técnicas, Conciliación Opción B y Generación de Plantillas.
  - Diagrama de secuencia Mermaid embebido.
- **`secuencia_etl.mmd`**:
  - Archivo fuente en notación Mermaid estándar para integración en herramientas de diagramado o documentación Markdown.

---

## 🚀 Cómo Re-renderizar o Actualizar con Archify

### En Windows (1-Clic):
```powershell
# PowerShell
.\scripts\diagramas\render_arquitectura.ps1

# O con apertura automática en navegador:
.\scripts\diagramas\render_arquitectura.ps1 -Open

# Símbolo del sistema (CMD / Batch):
scripts\diagramas\render_arquitectura.bat
```

### Directo con Archify CLI:
```bash
node C:\Users\mantenimiento\.agents\skills\archify\bin\archify.mjs render architecture scripts/diagramas/arquitectura.json scripts/diagramas/arquitectura.html
```

---

## 👁️ Visualización Directa

Para abrir los diagramas:
1. Haz doble clic en `arquitectura.html` o en `flujo_etl_neumaticos.html`.
2. Ábrelos con cualquier navegador moderno (Chrome, Edge, Firefox).
