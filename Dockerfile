FROM python:3.11-slim

# Metadatos del contenedor
LABEL maintainer="Transportes Díaz Acarreos Generales S.A.C."
LABEL description="Pipeline ETL Streamlit para DAG Neumáticos"

WORKDIR /app

# Instalar utilidades del sistema
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copiar requirements e instalar dependencias Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar scripts, configuración y aplicación
COPY scripts/ ./scripts/
COPY .streamlit/ ./.streamlit/
COPY app.py .

# Crear estructura de carpetas de datos
RUN mkdir -p Archivos/Maestro Archivos/Plantillas

# Exponer el puerto de Streamlit
EXPOSE 8502

# Healthcheck para verificar el estado de la app
HEALTHCHECK --interval=30s --timeout=10s --retries=3 \
    CMD curl --fail http://localhost:8502/_stcore/health || exit 1

# Comando de inicio de Streamlit
ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8502", "--server.address=0.0.0.0", "--server.headless=true", "--browser.gatherUsageStats=false"]
