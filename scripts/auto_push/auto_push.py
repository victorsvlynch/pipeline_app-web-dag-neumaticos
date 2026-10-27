"""
Auto-Push Python CLI: DAG Neumáticos
Transportes Díaz Acarreos Generales S.A.C.

Permite generar commits convencionales en español e integrarse con GitHub Actions.
"""

import subprocess
import os
import sys

def main():
    print("📦 Añadiendo cambios (git add .)...")
    subprocess.run(["git", "add", "."], check=True)
    
    # Obtener los cambios a commitear
    result = subprocess.run(["git", "diff", "--cached"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    diff = result.stdout.strip()
    
    if not diff:
        print("⚠️ No hay cambios nuevos para subir a GitHub.")
        return
        
    prompt = f"""Eres un desarrollador experto. Escribe un único mensaje de commit usando el estándar de "Conventional Commits" basado en el siguiente 'git diff'.

Tipos permitidos:
- feat: (para nuevas características, plantillas o vistas en Streamlit)
- fix: (para solución de errores en extracción o cálculos)
- refactor: (para mejoras o limpieza de código ETL)
- chore: (mantenimiento de dependencias, scripts o Docker)
- docs: (documentación o diagramas)

REGLAS ESTRICTAS:
1. Devuelve ÚNICAMENTE el mensaje del commit en una sola línea. No incluyas explicaciones, ni saludos, ni comillas, ni markdown.
2. El mensaje DEBE estar en idioma Español.
3. Sé conciso pero claro sobre lo que cambió.

DIFF DE CÓDIGO:
{diff[:30000]}
"""

    print("🧠 Copia el siguiente texto y pégalo en el chat de la IA para generar el mensaje de commit:\n")
    print("-" * 60)
    print(prompt)
    print("-" * 60)
    
    commit_msg = input("\n📝 Pega aquí el mensaje generado por la IA (o escribe el tuyo propio) y presiona Enter: ").strip()
    
    if not commit_msg:
        print("❌ Operación cancelada. No se ingresó ningún mensaje.")
        return
        
    # Limpiar comillas
    commit_msg = commit_msg.strip('`').strip('"').strip("'").replace('\n', ' ')
    
    try:
        print(f"\n🚀 Haciendo Commit y Push a GitHub con el mensaje: '{commit_msg}'...")
        subprocess.run(["git", "commit", "-m", commit_msg], check=True)
        try:
            subprocess.run(["git", "push", "origin", "main"], check=True)
        except subprocess.CalledProcessError:
            subprocess.run(["git", "push"], check=True)
        
        print("\n✅ ¡ÉXITO! El código fue subido a GitHub.")
        print("🤖 GitHub Actions (.github/workflows/release.yml) ahora detectará el prefijo y creará el Release SemVer automáticamente.")
        
    except Exception as e:
        print(f"❌ Error al subir a Git: {e}")

if __name__ == "__main__":
    main()
