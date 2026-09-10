$dir  = $PSScriptRoot
$dir2 = Join-Path (Split-Path $PSScriptRoot -Parent) "Fotos para catalogo HTLM en Fabrica"
$htmlPath = Join-Path $dir "Catalogo.html"
$entries  = [System.Collections.Generic.List[string]]::new()

# ── Carpeta principal: raíz ──────────────────────────────────────────────────
$jpgs = Get-ChildItem -Path $dir -Filter *.jpg | Sort-Object Name
foreach ($f in $jpgs) {
    $code = $f.BaseName.Replace('\', '\\').Replace('"', '\"')
    $name = $f.Name.Replace('\', '\\').Replace('"', '\"')
    $entries.Add('["' + $code + '", "' + $name + '"]')
}

# ── Carpeta principal: subcarpetas ───────────────────────────────────────────
$subfolders = Get-ChildItem -Path $dir -Directory | Sort-Object Name
foreach ($sub in $subfolders) {
    $subjpgs = Get-ChildItem -Path $sub.FullName -Filter *.jpg | Sort-Object Name
    foreach ($f in $subjpgs) {
        $code    = $f.BaseName.Replace('\', '\\').Replace('"', '\"')
        $subname = $sub.Name.Replace('\', '\\').Replace('"', '\"')
        $fname   = $f.Name.Replace('\', '\\').Replace('"', '\"')
        $entries.Add('["' + $code + '", "' + $subname + '/' + $fname + '"]')
    }
}

# ── Carpeta en fábrica: raíz ─────────────────────────────────────────────────
if (Test-Path $dir2) {
    $jpgs2 = Get-ChildItem -Path $dir2 -Filter *.jpg | Sort-Object Name
    foreach ($f in $jpgs2) {
        $code = $f.BaseName.Replace('\', '\\').Replace('"', '\"')
        $name = $f.Name.Replace('\', '\\').Replace('"', '\"')
        $entries.Add('["' + $code + '", "' + $name + '"]')
    }
    # ── Carpeta en fábrica: subcarpetas ──────────────────────────────────────
    $subfolders2 = Get-ChildItem -Path $dir2 -Directory | Sort-Object Name
    foreach ($sub in $subfolders2) {
        $subjpgs = Get-ChildItem -Path $sub.FullName -Filter *.jpg | Sort-Object Name
        foreach ($f in $subjpgs) {
            $code    = $f.BaseName.Replace('\', '\\').Replace('"', '\"')
            $subname = $sub.Name.Replace('\', '\\').Replace('"', '\"')
            $fname   = $f.Name.Replace('\', '\\').Replace('"', '\"')
            $entries.Add('["' + $code + '", "' + $subname + '/' + $fname + '"]')
        }
    }
} else {
    Write-Host "  (Carpeta en fabrica no encontrada, se omite)"
}

# ── Escribir items en Catalogo.html ─────────────────────────────────────────
$itemsLine = "const items = [" + ($entries -join ', ') + "];"
$content   = Get-Content -Path $htmlPath -Raw -Encoding UTF8
$pattern   = 'const items = \[.*?\];'
$evaluator = [System.Text.RegularExpressions.MatchEvaluator] { param($m) $itemsLine }
$newContent = [System.Text.RegularExpressions.Regex]::Replace(
    $content, $pattern, $evaluator,
    [System.Text.RegularExpressions.RegexOptions]::Singleline
)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($htmlPath, $newContent, $utf8NoBom)
Write-Host "Galeria actualizada: $($entries.Count) fotos."
