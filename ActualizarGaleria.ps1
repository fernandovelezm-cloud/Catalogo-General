# Regenera la lista de fotos (items) en Catalogo.html
# Incluye fotos en la carpeta raíz y en subcarpetas (un nivel).
$dir = $PSScriptRoot
$htmlPath = Join-Path $dir "Catalogo.html"

$entries = [System.Collections.Generic.List[string]]::new()

# Fotos en la carpeta raíz
$jpgs = Get-ChildItem -Path $dir -Filter *.jpg | Sort-Object Name
foreach ($f in $jpgs) {
    $code = $f.BaseName.Replace('\', '\\').Replace('"', '\"')
    $name = $f.Name.Replace('\', '\\').Replace('"', '\"')
    $entries.Add('["' + $code + '", "' + $name + '"]')
}

# Fotos en subcarpetas (un nivel)
$subfolders = Get-ChildItem -Path $dir -Directory | Sort-Object Name
foreach ($sub in $subfolders) {
    $subjpgs = Get-ChildItem -Path $sub.FullName -Filter *.jpg | Sort-Object Name
    foreach ($f in $subjpgs) {
        $code = $f.BaseName.Replace('\', '\\').Replace('"', '\"')
        $subname = $sub.Name.Replace('\', '\\').Replace('"', '\"')
        $filename = $f.Name.Replace('\', '\\').Replace('"', '\"')
        $entries.Add('["' + $code + '", "' + $subname + '/' + $filename + '"]')
    }
}

$itemsLine = "const items = [" + ($entries -join ', ') + "];"

$content = Get-Content -Path $htmlPath -Raw -Encoding UTF8
$pattern = 'const items = \[.*?\];'
$evaluator = [System.Text.RegularExpressions.MatchEvaluator] { param($m) $itemsLine }
$newContent = [System.Text.RegularExpressions.Regex]::Replace($content, $pattern, $evaluator, [System.Text.RegularExpressions.RegexOptions]::Singleline)

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($htmlPath, $newContent, $utf8NoBom)

Write-Host "Galeria actualizada: $($entries.Count) fotos."
