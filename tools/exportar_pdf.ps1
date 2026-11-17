# Exporta el .docx a PDF con Word en dos fases cortas:
#   1) abre el documento, actualiza la tabla de contenido y los campos, y guarda;
#   2) lo abre en solo lectura y exporta el PDF en una carpeta temporal, y lo copia a docs\.
# Motivo: en este equipo la licencia de Office pasa a "Unlicensed Product" unos segundos después de abrir Word, y en ese
# modo cualquier edición se queda esperando; además, Word rechaza exportar a una ruta con "+" (la del worktree).
param([string]$Nombre = 'Entrega2_IC_Proyecto', [int]$Limite = 180)
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$docx = Join-Path $raiz "docs\$Nombre.docx"
$pdf = Join-Path $raiz "docs\$Nombre.pdf"
$tmp = Join-Path $env:TEMP "exportar-$Nombre"
New-Item -ItemType Directory -Force $tmp | Out-Null

function Invoke-Fase([scriptblock]$Bloque, [object[]]$Argumentos) {
    $job = Start-Job -ScriptBlock $Bloque -ArgumentList $Argumentos
    if (-not (Wait-Job $job -Timeout $Limite)) {
        Stop-Job $job
        Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" | Where-Object { $_.CommandLine -match '/Automation' } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
        throw "Word no respondió en $Limite s (¿licencia de Office en 'Unlicensed Product'?)"
    }
    Receive-Job $job
}

Invoke-Fase {
    param($docx)
    $w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
    $d = $w.Documents.Open($docx)
    $d.Repaginate()
    if ($d.TablesOfContents.Count -gt 0) { $d.TablesOfContents.Item(1).Update() }
    $d.Fields.Update() | Out-Null
    $d.Repaginate()
    $d.Save()
    "Fase 1: tabla de contenido actualizada y guardada ($($w.ActiveWindow.Caption))"
    $d.Close([ref]$false); $w.Quit()
} @($docx)

$copia = Join-Path $tmp "$Nombre.docx"
Copy-Item $docx $copia -Force
Invoke-Fase {
    param($copia, $salida)
    $w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
    $d = $w.Documents.Open($copia, $false, $true)
    $d.ExportAsFixedFormat($salida, 17)
    "Fase 2: PDF exportado. Paginas: $($d.ComputeStatistics(2))"
    $d.Close([ref]$false); $w.Quit()
} @($copia, (Join-Path $tmp "$Nombre.pdf"))
Copy-Item (Join-Path $tmp "$Nombre.pdf") $pdf -Force
Write-Output "PDF: $pdf"
