param(
    [Parameter(Mandatory = $true)][string]$InputPdf,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [int]$Dpi = 144
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$genericAsTask = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq "AsTask" -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 } |
    Select-Object -First 1
$actionAsTask = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq "AsTask" -and -not $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 } |
    Select-Object -First 1

function Await-Operation {
    param($Operation, [Type]$ResultType)
    $method = $genericAsTask.MakeGenericMethod($ResultType)
    $task = $method.Invoke($null, @($Operation))
    $task.Wait()
    return $task.Result
}

function Await-Action {
    param($Action)
    $task = $actionAsTask.Invoke($null, @($Action))
    $task.Wait()
}

$pdfPath = (Resolve-Path -LiteralPath $InputPdf).Path
$outputDir = [System.IO.Path]::GetFullPath($OutputDirectory)
[void](New-Item -ItemType Directory -Force -Path $outputDir)

$storageFileType = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$pdfDocumentType = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$memoryStreamType = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$renderOptionsType = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]

$file = Await-Operation ($storageFileType::GetFileFromPathAsync($pdfPath)) $storageFileType
$pdf = Await-Operation ($pdfDocumentType::LoadFromFileAsync($file)) $pdfDocumentType

for ($i = 0; $i -lt $pdf.PageCount; $i++) {
    $page = $pdf.GetPage([uint32]$i)
    $stream = New-Object $memoryStreamType.FullName
    $options = New-Object $renderOptionsType.FullName
    $options.DestinationWidth = [uint32][Math]::Ceiling($page.Size.Width * ($Dpi / 96.0))
    $options.DestinationHeight = [uint32][Math]::Ceiling($page.Size.Height * ($Dpi / 96.0))
    Await-Action ($page.RenderToStreamAsync($stream, $options))
    $stream.Seek(0)
    $netStream = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream)
    $pngPath = Join-Path $outputDir ("page-{0:D2}.png" -f ($i + 1))
    $fileStream = [System.IO.File]::Open($pngPath, [System.IO.FileMode]::Create)
    try { $netStream.CopyTo($fileStream) } finally { $fileStream.Close(); $netStream.Close() }
    $stream.Dispose()
    $page.Dispose()
}

Write-Output ("Rendered {0} pages to {1}" -f $pdf.PageCount, $outputDir)
