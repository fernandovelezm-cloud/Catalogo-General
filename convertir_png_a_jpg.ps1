Add-Type -AssemblyName System.Drawing

$folder = Split-Path -Parent $MyInvocation.MyCommand.Path
$pngs = Get-ChildItem -Path $folder -Filter "*.png"

if ($pngs.Count -eq 0) {
    Write-Host "No se encontraron archivos PNG." -ForegroundColor Yellow
    exit
}

Write-Host "Encontrados $($pngs.Count) archivos PNG. Iniciando conversion..." -ForegroundColor Cyan

$converted = 0
$failed = 0

foreach ($png in $pngs) {
    $jpgPath = [System.IO.Path]::ChangeExtension($png.FullName, ".jpg")
    try {
        $img = [System.Drawing.Image]::FromFile($png.FullName)

        # Si tiene transparencia, componer sobre fondo blanco
        $bmp = New-Object System.Drawing.Bitmap($img.Width, $img.Height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.DrawImage($img, 0, 0, $img.Width, $img.Height)
        $g.Dispose()
        $img.Dispose()

        # Guardar como JPG con calidad 95
        $encoder = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq "image/jpeg" }
        $encParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
        $encParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 95L)
        $bmp.Save($jpgPath, $encoder, $encParams)
        $bmp.Dispose()

        Remove-Item $png.FullName -Force
        Write-Host "  Convertido: $($png.Name)" -ForegroundColor Green
        $converted++
    } catch {
        Write-Host "  Error en $($png.Name): $_" -ForegroundColor Red
        $failed++
    }
}

Write-Host ""
Write-Host "Completado: $converted convertidos, $failed errores." -ForegroundColor Cyan
Read-Host "Presiona Enter para cerrar"
