# Abre el .docx con Word, actualiza la tabla de contenido y los campos, guarda el .docx y exporta el PDF.
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$docx = Join-Path $raiz 'docs\Entrega1_IC_Proyecto1.docx'
$pdf  = Join-Path $raiz 'docs\Entrega1_IC_Proyecto1.pdf'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
try {
    $d = $w.Documents.Open($docx)
    $d.Repaginate()
    if ($d.TablesOfContents.Count -gt 0) { $d.TablesOfContents.Item(1).Update() }
    $d.Fields.Update() | Out-Null
    $d.Repaginate()
    $d.Save()
    $d.ExportAsFixedFormat($pdf, 17)
    $paginas = $d.ComputeStatistics(2)
    Write-Output "Paginas: $paginas"
    $d.Close([ref]$false)
}
finally {
    $w.Quit()
}
