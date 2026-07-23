param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputPdf
)

$word = $null
$document = $null
$wordProcessId = 0
$existingWordProcessIds = @()
try {
    $inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
    $outputPath = [System.IO.Path]::GetFullPath($OutputPdf)
    $outputDirectory = [System.IO.Path]::GetDirectoryName($outputPath)
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

    $existingWordProcessIds = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    if (-not ("WordWindowProcess" -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WordWindowProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
    }
    $wordHwnd = $word.Hwnd
    if ($null -ne $wordHwnd -and [int64]$wordHwnd -ne 0) {
        [uint32]$processId = 0
        [void][WordWindowProcess]::GetWindowThreadProcessId([IntPtr]$wordHwnd, [ref]$processId)
        $wordProcessId = [int]$processId
    }
    if ($wordProcessId -eq 0) {
        Start-Sleep -Milliseconds 750
        $newWordProcess = Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
            Where-Object { $existingWordProcessIds -notcontains $_.Id } |
            Sort-Object StartTime -Descending |
            Select-Object -First 1
        if ($null -ne $newWordProcess) {
            $wordProcessId = $newWordProcess.Id
        }
    }
    $document = $word.Documents.Open($inputPath, $false, $true)
    $document.ExportAsFixedFormat($outputPath, 17)
    Write-Output $outputPath
}
finally {
    if ($null -ne $document) {
        $document.Close([ref]$false)
    }
    if ($null -ne $word) {
        if ($wordProcessId -eq 0) {
            try { $word.Quit([ref]$false) } catch { }
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    if ($wordProcessId -gt 0) {
        Stop-Process -Id $wordProcessId -Force -ErrorAction SilentlyContinue
    }
}
