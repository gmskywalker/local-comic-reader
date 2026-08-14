[CmdletBinding()]
param(
    [string]$OutputDirectory = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $PSScriptRoot 'single-file-tools'
}

function ConvertTo-GzipBase64 {
    param([string]$Text)
    $inputBytes = [Text.Encoding]::UTF8.GetBytes($Text)
    $output = New-Object IO.MemoryStream
    try {
        $gzip = New-Object IO.Compression.GZipStream($output, [IO.Compression.CompressionMode]::Compress, $true)
        try { $gzip.Write($inputBytes, 0, $inputBytes.Length) }
        finally { $gzip.Dispose() }
        return [Convert]::ToBase64String($output.ToArray())
    }
    finally { $output.Dispose() }
}

function New-HiddenVbsTool {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$DisplayName
    )
    $source = [IO.File]::ReadAllText($SourcePath, [Text.Encoding]::UTF8)
    $payload = ConvertTo-GzipBase64 -Text $source
    $payloadLines = New-Object 'System.Collections.Generic.List[string]'
    for ($index = 0; $index -lt $payload.Length; $index += 120) {
        $length = [Math]::Min(120, $payload.Length - $index)
        $payloadLines.Add("'" + $payload.Substring($index, $length))
    }

    $loader = @'
try {
    $raw = [IO.File]::ReadAllText($env:LOCAL_COMIC_TOOL_PATH, [Text.Encoding]::UTF8)
    $marker = "'#==POWERSHELL_PAYLOAD=="
    $pos = $raw.LastIndexOf($marker)
    if ($pos -lt 0) { throw '找不到内嵌程序' }
    $encoded = $raw.Substring($pos + $marker.Length) -replace '[^A-Za-z0-9+/=]', ''
    $bytes = [Convert]::FromBase64String($encoded)
    $inputStream = New-Object IO.MemoryStream(,$bytes)
    $gzip = New-Object IO.Compression.GZipStream($inputStream, [IO.Compression.CompressionMode]::Decompress)
    $reader = New-Object IO.StreamReader($gzip, [Text.Encoding]::UTF8)
    $source = $reader.ReadToEnd()
    $reader.Dispose()
    $gzip.Dispose()
    $inputStream.Dispose()
    $rawArgs = @()
    if (-not [string]::IsNullOrEmpty($env:LOCAL_COMIC_TOOL_ARGUMENTS)) {
        $rawArgs = @($env:LOCAL_COMIC_TOOL_ARGUMENTS -split [char]30)
    }
    $toolParams = @{}
    for ($index = 0; $index -lt $rawArgs.Count; $index++) {
        $name = $rawArgs[$index].TrimStart('-')
        if ($index + 1 -lt $rawArgs.Count -and $rawArgs[$index + 1] -notmatch '^-') {
            $toolParams[$name] = $rawArgs[$index + 1]
            $index++
        }
        else {
            $toolParams[$name] = $true
        }
    }
    $block = [scriptblock]::Create($source)
    & $block @toolParams
}
catch {
    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Forms.MessageBox]::Show($_.Exception.Message, '__DISPLAY_NAME__', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit 1
}
'@
    $loader = $loader.Replace('__DISPLAY_NAME__', $DisplayName.Replace("'", "''"))
    $loaderEncoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($loader))

    $settingsLine = "'#==TOOL_SETTINGS==e30="
    if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) {
        try {
            $existingContent = [IO.File]::ReadAllText($DestinationPath, [Text.Encoding]::UTF8)
            $settingsMatches = [regex]::Matches($existingContent, '(?m)^''#==TOOL_SETTINGS==(?<data>[A-Za-z0-9+/=]*)(?=\r?$)')
            if ($settingsMatches.Count -eq 1) {
                $settingsJson = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($settingsMatches[0].Groups['data'].Value))
                [void]($settingsJson | ConvertFrom-Json)
                $settingsLine = $settingsMatches[0].Value
            }
        }
        catch {
            # 旧文件没有有效嵌入设置时，从空设置开始；程序正文仍照常重新打包。
        }
    }

    $header = @"
Option Explicit
Dim shell, processEnv, selfPath, argumentData, item, commandLine, exitCode
Set shell = CreateObject("WScript.Shell")
Set processEnv = shell.Environment("Process")
selfPath = WScript.ScriptFullName
argumentData = ""
For Each item In WScript.Arguments
    If Len(argumentData) > 0 Then argumentData = argumentData & Chr(30)
    argumentData = argumentData & CStr(item)
Next
processEnv("LOCAL_COMIC_TOOL_PATH") = selfPath
processEnv("LOCAL_COMIC_TOOL_ARGUMENTS") = argumentData
commandLine = "powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand $loaderEncoded"
exitCode = shell.Run(commandLine, 0, True)
WScript.Quit exitCode
$settingsLine
'#==POWERSHELL_PAYLOAD==
"@
    $content = $header + ($payloadLines -join "`r`n") + "`r`n"
    try {
        [IO.File]::WriteAllText($DestinationPath, $content, $utf8NoBom)
    }
    catch {
        throw ("Cannot write packaged file [$DestinationPath]: " + $_.Exception.Message)
    }
}

if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    [void](New-Item -ItemType Directory -Path $OutputDirectory)
}
New-HiddenVbsTool -SourcePath (Join-Path $PSScriptRoot 'comic-reader-generator.ps1') -DestinationPath (Join-Path $OutputDirectory '漫画更新器.vbs') -DisplayName '漫画更新器错误'
New-HiddenVbsTool -SourcePath (Join-Path $PSScriptRoot 'comic-organizer.ps1') -DestinationPath (Join-Path $OutputDirectory '漫画整理器.vbs') -DisplayName '漫画整理器错误'
$cbzSourcePath = Join-Path $PSScriptRoot 'comic-cbz-exporter.ps1'
if (Test-Path -LiteralPath $cbzSourcePath -PathType Leaf) {
    New-HiddenVbsTool -SourcePath $cbzSourcePath -DestinationPath (Join-Path $OutputDirectory '漫画导出器.vbs') -DisplayName '漫画 CBZ / EPUB / PDF 导出器错误'
}

Get-ChildItem -LiteralPath $OutputDirectory -File | Select-Object Name, Length
