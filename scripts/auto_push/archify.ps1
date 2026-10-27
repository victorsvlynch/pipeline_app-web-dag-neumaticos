<#
.SYNOPSIS
    Archify: Herramienta de empaquetado, respaldos de plantillas y snapshots del proyecto.
.DESCRIPTION
    Empaqueta y respalda las 7 plantillas oficiales Excel generadas por el pipeline ETL,
    los reportes de auditoria y conciliacion, o genera snapshots completos y limpios del codigo fuente.
.PARAMETER Mode
    Modo de empaquetado:
    - Plantillas: Respalda las 7 plantillas Excel y reportes de auditoria de /Archivos/Plantillas y /Archivos/Reportes.
    - Full: Genera snapshot limpio del codigo fuente excluyendo .git, .venv, __pycache__, archives y temporales.
    - All: Combina snapshot de codigo y respaldo de plantillas en archivos independientes.
.PARAMETER Tag
    Etiqueta personalizada para el nombre del archivo ZIP.
.PARAMETER OutputDir
    Directorio de destino (por defecto: <repo_root>/archives).
.PARAMETER CleanAfterArchive
    Si se especifica, limpia archivos temporales tras generar el ZIP.
.PARAMETER DryRun
    Modo simulacion, muestra los archivos que se incluirian sin crear el archivo ZIP.
.EXAMPLE
    .\scripts\auto_push\archify.ps1 -Mode Plantillas
.EXAMPLE
    .\scripts\auto_push\archify.ps1 -Mode Full -Tag "v1.2.0"
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, Mandatory = $false)]
    [ValidateSet("Plantillas", "Full", "All")]
    [string]$Mode = "Plantillas",

    [Parameter(Mandatory = $false)]
    [Alias("t", "Name")]
    [string]$Tag = "",

    [Parameter(Mandatory = $false)]
    [Alias("o")]
    [string]$OutputDir = "",

    [Parameter(Mandatory = $false)]
    [switch]$CleanAfterArchive,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Determinar raíz del proyecto
try {
    $repoRoot = (git rev-parse --show-toplevel 2>$null)
    if ($LASTEXITCODE -eq 0 -and $repoRoot -and (Test-Path $repoRoot.Trim())) {
        $repoRoot = $repoRoot.Trim()
    } else {
        $repoRoot = (Resolve-Path "$PSScriptRoot\..\..").Path
    }
} catch {
    $repoRoot = (Resolve-Path "$PSScriptRoot\..\..").Path
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $repoRoot "archives"
}
if (-not (Test-Path $OutputDir) -and -not $DryRun) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$gitBranch = "main"
try {
    $bOut = & git -C $repoRoot branch --show-current 2>&1
    if ($LASTEXITCODE -eq 0 -and $bOut) { $gitBranch = $bOut.ToString().Trim() }
} catch { $gitBranch = "main" }

$gitCommit = "initial"
try {
    $cOut = & git -C $repoRoot rev-parse --short HEAD 2>&1
    if ($LASTEXITCODE -eq 0 -and $cOut) { $gitCommit = $cOut.ToString().Trim() }
} catch { $gitCommit = "initial" }

$tagSuffix = if ($Tag) { "_$($Tag -replace '[^a-zA-Z0-9_\-]', '_')" } else { "" }

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " 📦 ARCHIFY: Empaquetado y Respaldos (DAG Neumáticos ETL)" -ForegroundColor Cyan
Write-Host " Transportes Díaz Acarreos Generales S.A.C." -ForegroundColor Gray
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "📁 Directorio base:  $repoRoot" -ForegroundColor Gray
Write-Host "⚙️  Modo activo:      $Mode" -ForegroundColor White
Write-Host "🌿 Rama / Commit:    $gitBranch ($gitCommit)" -ForegroundColor Gray
Write-Host "🎯 Destino:          $OutputDir" -ForegroundColor Gray
if ($DryRun) {
    Write-Host "🔍 MODO DRY-RUN ACTIVO (Simulacion sin escritura)" -ForegroundColor Magenta
}
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

function Create-ArchiveFromFiles {
    param(
        [string]$ZipPath,
        [System.Collections.Generic.List[PSObject]]$FileList,
        [hashtable]$ManifestData,
        [switch]$DeleteSources
    )

    if ($FileList.Count -eq 0) {
        Write-Host "ℹ️  No se encontraron archivos para incluir en el archivo ZIP." -ForegroundColor Yellow
        return
    }

    $totalBytes = 0
    foreach ($item in $FileList) { $totalBytes += $item.Length }

    Write-Host "`n📊 Archivos a empaquetar: $($FileList.Count) ($([math]::Round($totalBytes / 1MB, 2)) MB)" -ForegroundColor Green

    if ($DryRun) {
        Write-Host "`n📋 Vista previa (primeros 20 archivos):" -ForegroundColor Cyan
        $FileList | Select-Object -First 20 | ForEach-Object {
            Write-Host "   + $($_.RelativePath) ($([math]::Round($_.Length / 1KB, 1)) KB)" -ForegroundColor DarkGray
        }
        return
    }

    $manifestPath = Join-Path $env:TEMP "manifest_$timestamp.json"
    $ManifestData.filesCount = $FileList.Count
    $ManifestData.uncompressedBytes = $totalBytes
    $ManifestData.createdAt = (Get-Date).ToString("o")
    $ManifestData | ConvertTo-Json -Depth 5 | Out-File -FilePath $manifestPath -Encoding utf8

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression

    if (Test-Path $ZipPath) { Remove-Item -Path $ZipPath -Force }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $zip = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Create)

    try {
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip, $manifestPath, "manifest.json", [System.IO.Compression.CompressionLevel]::Optimal
        ) | Out-Null

        foreach ($file in $FileList) {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $zip, $file.FullPath, $file.RelativePath.Replace('\', '/'), [System.IO.Compression.CompressionLevel]::Optimal
            ) | Out-Null
        }
    } finally {
        $zip.Dispose()
        if (Test-Path $manifestPath) { Remove-Item -Path $manifestPath -Force -ErrorAction SilentlyContinue }
    }

    $sw.Stop()
    $zipSize = (Get-Item $ZipPath).Length
    $ratio = if ($totalBytes -gt 0) { [math]::Round((1 - ($zipSize / $totalBytes)) * 100, 1) } else { 0 }

    Write-Host "`n============================================================" -ForegroundColor Green
    Write-Host " 🎉 ARCHIVO GENERADO CON ÉXITO" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "📁 Archivo:       $ZipPath" -ForegroundColor White
    Write-Host "📄 Archivos:      $($FileList.Count) empaquetados (+ manifest.json)" -ForegroundColor White
    Write-Host "💾 Tamaño orig:   $([math]::Round($totalBytes / 1MB, 2)) MB" -ForegroundColor White
    Write-Host "🗜️  Tamaño ZIP:    $([math]::Round($zipSize / 1MB, 2)) MB (Reduccion: $ratio%)" -ForegroundColor White
    Write-Host "⏱️  Tiempo:        $([math]::Round($sw.Elapsed.TotalSeconds, 2)) segundos" -ForegroundColor White
}

if ($Mode -eq "Plantillas" -or $Mode -eq "All") {
    Write-Host "`n🔍 Buscando plantillas oficiales generadas y reportes de auditoria..." -ForegroundColor Yellow
    $zipName = "plantillas_dag_neumaticos_${timestamp}_${gitBranch}${tagSuffix}.zip"
    $zipPath = Join-Path $OutputDir $zipName

    $plantillaFiles = [System.Collections.Generic.List[PSObject]]::new()
    $searchDirs = @(
        (Join-Path $repoRoot "Archivos\Plantillas"),
        (Join-Path $repoRoot "Archivos\Reportes")
    )

    foreach ($dir in $searchDirs) {
        if (Test-Path $dir) {
            $files = Get-ChildItem -Path $dir -File | Where-Object {
                $_.Extension -match "\.(xlsx|zip)$" -and
                -not ($_.Name.StartsWith("~$"))
            }
            foreach ($f in $files) {
                $rel = $f.FullName.Substring($repoRoot.Length).TrimStart('\', '/')
                $plantillaFiles.Add([PSCustomObject]@{
                    FullPath     = $f.FullName
                    RelativePath = $rel
                    Length       = $f.Length
                })
            }
        }
    }

    $manifest = @{
        type        = "PlantillasYReportesBackup"
        project     = "Pipeline_App-Web-DAG-Neumaticos"
        company     = "Transportes Díaz Acarreos Generales S.A.C."
        timestamp   = $timestamp
        gitBranch   = $gitBranch
        gitCommit   = $gitCommit
    }

    Create-ArchiveFromFiles -ZipPath $zipPath -FileList $plantillaFiles -ManifestData $manifest
}

if ($Mode -eq "Full" -or $Mode -eq "All") {
    Write-Host "`n🔍 Escaneando codigo fuente del proyecto para Snapshot..." -ForegroundColor Yellow
    $zipName = "snapshot_source_${timestamp}_${gitBranch}_${gitCommit}${tagSuffix}.zip"
    $zipPath = Join-Path $OutputDir $zipName

    $excludeDirs = @(".git", ".venv", "venv", "env", "__pycache__", "archives", ".vscode", ".idea")
    $excludeFiles = @(".DS_Store", "Thumbs.db", "desktop.ini")

    $sourceFiles = [System.Collections.Generic.List[PSObject]]::new()
    $all = Get-ChildItem -Path $repoRoot -Recurse -File -Force

    foreach ($f in $all) {
        $rel = $f.FullName.Substring($repoRoot.Length).TrimStart('\', '/')
        $parts = $rel.Split([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)

        $skip = $false
        foreach ($p in $parts) {
            if ($excludeDirs -contains $p) { $skip = $true; break }
        }
        if ($skip) { continue }
        if ($excludeFiles -contains $f.Name) { continue }
        if ($f.Name.StartsWith("~$")) { continue }
        if ($f.Name.StartsWith(".env") -and -not $f.Name.EndsWith(".example")) { continue }
        if ($f.Extension -match "\.(zip|tar|gz|rar)$") { continue }

        $sourceFiles.Add([PSCustomObject]@{
            FullPath     = $f.FullName
            RelativePath = $rel
            Length       = $f.Length
        })
    }

    $manifest = @{
        type        = "FullSourceSnapshot"
        project     = "Pipeline_App-Web-DAG-Neumaticos"
        company     = "Transportes Díaz Acarreos Generales S.A.C."
        timestamp   = $timestamp
        gitBranch   = $gitBranch
        gitCommit   = $gitCommit
        tag         = $Tag
    }

    Create-ArchiveFromFiles -ZipPath $zipPath -FileList $sourceFiles -ManifestData $manifest
}
