$ErrorActionPreference = "Stop"

# Determinar raíz del repositorio
Push-Location "$PSScriptRoot\..\.."

try {
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " 🚀 AUTO-PUSH: Pipeline ETL DAG Neumáticos (Streamlit)" -ForegroundColor Cyan
    Write-Host " Transportes Díaz Acarreos Generales S.A.C." -ForegroundColor Gray
    Write-Host "============================================================" -ForegroundColor Cyan

    Write-Host "Añadiendo cambios (git add .)..." -ForegroundColor Yellow
    git add .

    # Obtener los cambios a commitear usando codificacion UTF-8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $diff = (git diff --cached | Out-String)

    if ([string]::IsNullOrWhiteSpace($diff)) {
        Write-Host "No hay cambios nuevos para subir a GitHub." -ForegroundColor Green
        return
    }

    # Limitar el tamaño del diff para no exceder la memoria del prompt
    if ($diff.Length -gt 25000) {
        $diff = $diff.Substring(0, 25000)
    }

    # Verificar si el usuario paso un parametro de mensaje
    $commitMsg = ""
    if ($args.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($args[0])) {
        $commitMsg = $args[0].ToString().Trim()
    }

    if ([string]::IsNullOrWhiteSpace($commitMsg)) {
        Write-Host "Analizando cambios y generando descripcion con IA (Gemini)..." -ForegroundColor Magenta

        # Buscar API Key en api_keys.txt o variable de entorno
        $apiKey = ""
        $keyPath = "$PSScriptRoot\api_keys.txt"
        if (Test-Path $keyPath) {
            $lines = Get-Content $keyPath
            foreach ($line in $lines) {
                $cleanKey = $line.Trim().Replace('"', '').Replace("'", "")
                if ($cleanKey.Length -gt 20 -and -not $cleanKey.StartsWith("#")) {
                    $apiKey = $cleanKey
                    break
                }
            }
        }
        if ([string]::IsNullOrWhiteSpace($apiKey) -and $env:GEMINI_API_KEY) {
            $apiKey = $env:GEMINI_API_KEY.Trim()
        }

        if ([string]::IsNullOrWhiteSpace($apiKey)) {
            Write-Host "Aviso: No se encontro una API Key valida en api_keys.txt o GEMINI_API_KEY" -ForegroundColor Yellow
            $nowStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $commitMsg = "feat: actualizacion de pipeline ETL y dashboard Streamlit ($nowStr)"
        } else {
            $promptLines = @(
                "Eres un desarrollador experto trabajando en el proyecto 'Pipeline ETL y Dashboard Streamlit - DAG Neumaticos' para Transportes Diaz Acarreos Generales S.A.C.",
                "Escribe un unico mensaje de commit usando el estandar de 'Conventional Commits' basado en el siguiente 'git diff'.",
                "",
                "Tipos permitidos:",
                "- feat: (para nuevas caracteristicas, plantillas o vistas en Streamlit)",
                "- fix: (para solucion de errores en extraccion o calculos)",
                "- refactor: (para mejoras o limpieza de codigo ETL)",
                "- chore: (mantenimiento de dependencias, scripts de despliegue o Docker)",
                "- docs: (documentacion o diagramas)",
                "",
                "REGLAS ESTRICTAS:",
                "1. Devuelve UNICAMENTE el mensaje del commit en una sola linea. No incluyas explicaciones, saludos, comillas ni formato markdown.",
                "2. El mensaje DEBE estar en idioma Espanol.",
                "3. Se conciso pero claro sobre lo que cambio tecnicamente.",
                "",
                "DIFF DE CODIGO:",
                $diff
            )
            $prompt = $promptLines -join [Environment]::NewLine

            $bodyObj = @{
                contents = @(
                    @{
                        parts = @(
                            @{
                                text = $prompt
                            }
                        )
                    }
                )
            }
            $bodyJson = $bodyObj | ConvertTo-Json -Depth 10

            # Probar modelos disponibles con fallback automatico
            $models = @("gemini-2.5-flash-lite", "gemini-2.5-flash", "gemini-flash-latest")
            $success = $false
            $response = $null

            foreach ($modelName in $models) {
                $url = "https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=$apiKey"
                try {
                    $response = Invoke-RestMethod -Uri $url -Method Post -Body $bodyJson -ContentType "application/json; charset=utf-8"
                    if ($response.candidates -and $response.candidates[0].content.parts) {
                        $success = $true
                        break
                    }
                } catch {
                    $statusCode = 0
                    if ($_.Exception.Response) {
                        $statusCode = [int]$_.Exception.Response.StatusCode
                    }
                    if ($statusCode -eq 503 -or $statusCode -eq 429) {
                        Write-Host "   Modelo $modelName saturado ($statusCode). Probando alternativa..." -ForegroundColor Yellow
                        continue
                    }
                }
            }

            if ($success) {
                # Extraer y limpiar el mensaje
                $commitMsg = $response.candidates[0].content.parts[0].text.Trim()
                $commitMsg = $commitMsg.Replace([char]34, ' ').Replace([char]39, ' ')
                $commitMsg = $commitMsg.Replace([Environment]::NewLine, " ").Trim()
            } else {
                $nowStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $commitMsg = "feat: actualizacion y sincronizacion del proyecto ($nowStr)"
            }
        }
    }

    Write-Host ""
    Write-Host "Mensaje generado:" -ForegroundColor Cyan
    Write-Host "> $commitMsg" -ForegroundColor White
    Write-Host ""

    Write-Host "Haciendo Commit y Push a GitHub..." -ForegroundColor Yellow
    git commit -m "$commitMsg"
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $pushRes = & git push origin main 2>&1
    if ($LASTEXITCODE -ne 0) {
        $pushRes = & git push origin master 2>&1
        if ($LASTEXITCODE -ne 0) {
            $pushRes = & git push 2>&1
        }
    }
    $ErrorActionPreference = $prevEAP

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host " 🎉 ¡EXITO! El codigo fue sincronizado con GitHub." -ForegroundColor Green
    Write-Host " 🤖 GitHub Actions (.github/workflows/release.yml) ahora" -ForegroundColor Cyan
    Write-Host "    detectara el prefijo y creara el Release SemVer y Tag" -ForegroundColor Cyan
    Write-Host "    automaticamente en GitHub." -ForegroundColor Cyan
    Write-Host " 💡 Para sincronizar el nuevo Tag generado en tu PC:" -ForegroundColor DarkGray
    Write-Host "    git fetch --tags" -ForegroundColor DarkGray
    Write-Host "============================================================" -ForegroundColor Green

} catch {
    Write-Host "Error al comunicarse con la API o ejecutar git:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
} finally {
    Pop-Location
}
