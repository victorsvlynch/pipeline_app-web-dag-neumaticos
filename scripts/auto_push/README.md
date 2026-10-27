# 🚀 Auto-Push & Archify: Automatización Git y Respaldos

Herramientas de sincronización inteligente con **GitHub Actions (SemVer)** y empaquetado de respaldos y snapshots con **Archify**.

---

## 📂 Contenido del Directorio

| Archivo | Descripción |
| :--- | :--- |
| **`auto_push.ps1`** | Script de PowerShell que analiza el `git diff`, genera automáticamente el commit convencional en español utilizando **Google Gemini AI** (o fallback con timestamp) y realiza push a la rama `main`. |
| **`auto_push.bat`** | Lanzador en 1-clic para Windows que ejecuta `auto_push.ps1` o solicita un mensaje manual. |
| **`auto_push.py`** | Alternativa en Python interactivo para terminal. |
| **`archify.ps1`** | Herramienta de empaquetado en formato ZIP (.NET nativo) con manifiesto JSON estructurado (`manifest.json`). |
| **`archify.bat`** | Lanzador en 1-clic de `archify.ps1` con selección de modo interactiva. |
| **`api_keys.txt`** | Archivo de configuración local donde puedes ingresar tu API Key de Google Gemini para la generación de commits. |

---

## 🤖 Flujo de Versiones Semánticas (GitHub Actions)

Al hacer `git push` a la rama `main`, el workflow de GitHub Actions [`.github/workflows/release.yml`](../../.github/workflows/release.yml) lee el prefijo del commit para incrementar la versión automáticamente:

- **`fix:`** ➡️ Incrementa la versión PATCH (ej. `v1.0.0` ➡️ `v1.0.1`).
- **`feat:`** ➡️ Incrementa la versión MINOR (ej. `v1.0.1` ➡️ `v1.1.0`).
- **`feat!:` o `BREAKING CHANGE:`** ➡️ Incrementa la versión MAJOR (ej. `v1.1.0` ➡️ `v2.0.0`).

Automáticamente se crea un **GitHub Tag** y un **GitHub Release** con el changelog detallado.

---

## 📦 Uso de Archify (Empaquetado y Snapshots)

### Modos Disponibles:
1. **`Plantillas` (Por defecto):** Respalda las 7 plantillas Excel oficiales generadas por el ETL y los reportes de auditoría en la carpeta `archives/`.
2. **`Full`:** Genera un snapshot completo y limpio del código fuente del proyecto listo para distribución.
3. **`All`:** Genera ambos archivos ZIP de forma independiente.

### Ejemplos en PowerShell:
```powershell
# Respaldar plantillas y reportes generados
.\scripts\auto_push\archify.ps1 -Mode Plantillas

# Generar snapshot de código con etiqueta personalizada
.\scripts\auto_push\archify.ps1 -Mode Full -Tag "v1.1.0"

# Modo simulación (Dry-Run sin escribir)
.\scripts\auto_push\archify.ps1 -Mode All -DryRun
```

### Ejecución en 1-Clic:
Haga doble clic en **`scripts\auto_push\archify.bat`**.
