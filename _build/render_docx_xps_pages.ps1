param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [int]$Dpi = 144
)

$ErrorActionPreference = "Stop"
$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
$outputDir = [System.IO.Path]::GetFullPath($OutputDirectory)
[void](New-Item -ItemType Directory -Force -Path $outputDir)
$xpsPath = Join-Path $outputDir "lesson-preview.xps"

$word = $null
$document = $null
$wordProcessId = 0
$existingWordProcessIds = @()

try {
    $existingWordProcessIds = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    if (-not ("WordWindowProcessRenderXps" -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WordWindowProcessRenderXps {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
    }
    if ($null -ne $word.Hwnd -and [int64]$word.Hwnd -ne 0) {
        [uint32]$processId = 0
        [void][WordWindowProcessRenderXps]::GetWindowThreadProcessId([IntPtr]$word.Hwnd, [ref]$processId)
        $wordProcessId = [int]$processId
    }
    if ($wordProcessId -eq 0) {
        Start-Sleep -Milliseconds 500
        $newWordProcess = Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
            Where-Object { $existingWordProcessIds -notcontains $_.Id } |
            Sort-Object StartTime -Descending |
            Select-Object -First 1
        if ($null -ne $newWordProcess) { $wordProcessId = $newWordProcess.Id }
    }
    $document = $word.Documents.Open($inputPath, $false, $true)
    # wdExportFormatXPS = 1
    $document.ExportAsFixedFormat($xpsPath, 1)
}
finally {
    if ($null -ne $document) {
        try { $document.Close([ref]$false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit([ref]$false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($wordProcessId -gt 0) { Stop-Process -Id $wordProcessId -Force -ErrorAction SilentlyContinue }
}

Add-Type -AssemblyName ReachFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$xps = New-Object System.Windows.Xps.Packaging.XpsDocument($xpsPath, [System.IO.FileAccess]::Read)
try {
    $sequence = $xps.GetFixedDocumentSequence()
    $paginator = $sequence.DocumentPaginator
    $scale = $Dpi / 96.0
    for ($i = 0; $i -lt $paginator.PageCount; $i++) {
        $page = $paginator.GetPage($i)
        $pixelWidth = [Math]::Ceiling($page.Size.Width * $scale)
        $pixelHeight = [Math]::Ceiling($page.Size.Height * $scale)
        $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap(
            $pixelWidth,
            $pixelHeight,
            $Dpi,
            $Dpi,
            [System.Windows.Media.PixelFormats]::Pbgra32
        )
        $bitmap.Render($page.Visual)
        $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        [void]$encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $pngPath = Join-Path $outputDir ("page-{0:D2}.png" -f ($i + 1))
        $stream = [System.IO.File]::Open($pngPath, [System.IO.FileMode]::Create)
        try { $encoder.Save($stream) } finally { $stream.Close() }
    }
    Write-Output ("Rendered {0} pages to {1}" -f $paginator.PageCount, $outputDir)
}
finally {
    $xps.Close()
}
