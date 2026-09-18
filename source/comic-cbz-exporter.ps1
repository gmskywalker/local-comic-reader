[CmdletBinding()]
param(
    [string]$RootPath = '',
    [string]$OutputPath = '',
    [string]$ComicName = '',
    [ValidateSet('PerChapter', 'SingleBook', 'Epub', 'PdfStrip')]
    [string]$Mode = 'PerChapter',
    [switch]$IncludeCover,
    [switch]$AppendFormatToFolderName,
    [switch]$AutoNumberDuplicates,
    [switch]$SplitRootGroups,
    [switch]$NonInteractive,
    [switch]$ValidateOnly,
    [switch]$ForceIssues,
    [switch]$UiSmokeTest,
    [switch]$SkipOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ImageExtensions = @('.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.avif')
$script:ReaderResourceFolderName = '漫画阅读器资源'
$script:DefaultOutputFolderName = 'CBZ导出'
$script:ToolSettingsMarker = "'#==TOOL_SETTINGS=="
$script:ExporterToolPath = ''
$script:LegacyExporterSettingsRegistryPath = 'Software\LocalComicTools\ComicExporter'
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:ExporterProgressCallback = $null
$script:ExporterIsRunning = $false

function Initialize-ComicToolSharpText {
    if ($null -eq ('LocalComicSharpText.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace LocalComicSharpText {
    public static class NativeMethods {
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
    }
}
'@
    }
    try {
        [void][LocalComicSharpText.NativeMethods]::SetProcessDpiAwarenessContext([IntPtr](-4))
    }
    catch {
        # Older Windows versions continue with the original rendering mode.
    }
    Add-Type -AssemblyName System.Windows.Forms
    try {
        [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
    }
    catch [System.InvalidOperationException] {
        # A hosting process may already have created a WinForms window.
    }
}

function Get-ComicToolVisualControlTree {
    param([object]$Root)

    foreach ($child in @($Root.Controls)) {
        $child
        foreach ($descendant in @(Get-ComicToolVisualControlTree -Root $child)) {
            $descendant
        }
    }
}

function Set-ComicToolNativeDpiLayout {
    param([object]$Window)

    if ($null -eq $Window -or $null -ne $Window.PSObject.Properties['ComicToolNativeDpiApplied']) { return }
    Add-Member -InputObject $Window -MemberType NoteProperty -Name ComicToolNativeDpiApplied -Value $true

    $graphics = $Window.CreateGraphics()
    try { $scale = [double]$graphics.DpiX / 96.0 }
    finally { $graphics.Dispose() }
    if ($scale -le 1.001) { return }

    $grids = @()
    $lists = @()
    foreach ($control in @(Get-ComicToolVisualControlTree -Root $Window)) {
        if ($control -is [System.Windows.Forms.DataGridView]) {
            $grids += [pscustomobject]@{
                Control = $control
                HeaderHeight = [int]$control.ColumnHeadersHeight
                Columns = @($control.Columns | ForEach-Object {
                    [pscustomobject]@{ Column = $_; Width = [int]$_.Width; MinimumWidth = [int]$_.MinimumWidth }
                })
            }
        }
        elseif ($control -is [System.Windows.Forms.ListView]) {
            $lists += [pscustomobject]@{
                Control = $control
                Columns = @($control.Columns | ForEach-Object {
                    [pscustomobject]@{ Column = $_; Width = [int]$_.Width }
                })
            }
        }
    }

    $Window.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $Window.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $Window.PerformAutoScale()

    foreach ($gridState in $grids) {
        $grid = $gridState.Control
        if ($grid.ColumnHeadersHeightSizeMode -ne [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize) {
            $grid.ColumnHeadersHeight = [Math]::Max(4, [int][Math]::Round($gridState.HeaderHeight * $scale))
        }
        $grid.RowTemplate.Height = [Math]::Max($grid.RowTemplate.MinimumHeight, [int][Math]::Round(23 * $scale))
        foreach ($row in $grid.Rows) { $row.Height = $grid.RowTemplate.Height }
        foreach ($columnState in $gridState.Columns) {
            $columnState.Column.MinimumWidth = [Math]::Max(2, [int][Math]::Round($columnState.MinimumWidth * $scale))
            $columnState.Column.Width = [Math]::Max($columnState.Column.MinimumWidth, [int][Math]::Round($columnState.Width * $scale))
        }
    }
    foreach ($listState in $lists) {
        foreach ($columnState in $listState.Columns) {
            $columnState.Column.Width = [Math]::Max(1, [int][Math]::Round($columnState.Width * $scale))
        }
    }
}

function Set-ComicToolVisualTheme {
    param([object]$Window)

    if ($null -eq $Window) { return }
    Set-ComicToolNativeDpiLayout -Window $Window

    $canvas = [System.Drawing.Color]::FromArgb(246, 248, 251)
    $surface = [System.Drawing.Color]::White
    $surfaceAlt = [System.Drawing.Color]::FromArgb(241, 245, 249)
    $text = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $muted = [System.Drawing.Color]::FromArgb(91, 103, 119)
    $border = [System.Drawing.Color]::FromArgb(203, 213, 225)
    $accent = [System.Drawing.Color]::FromArgb(29, 111, 193)
    $accentHover = [System.Drawing.Color]::FromArgb(23, 92, 161)
    $accentPressed = [System.Drawing.Color]::FromArgb(18, 72, 126)
    $accentSoft = [System.Drawing.Color]::FromArgb(232, 243, 252)
    $selection = [System.Drawing.Color]::FromArgb(219, 237, 252)
    $selectionText = [System.Drawing.Color]::FromArgb(22, 55, 82)
    $danger = [System.Drawing.Color]::FromArgb(177, 35, 42)
    $dangerSoft = [System.Drawing.Color]::FromArgb(255, 239, 240)
    $gridLine = [System.Drawing.Color]::FromArgb(226, 232, 240)
    $defaultButtonBack = [System.Drawing.SystemColors]::Control.ToArgb()

    $Window.BackColor = $canvas
    $Window.ForeColor = $text

    foreach ($control in @(Get-ComicToolVisualControlTree -Root $Window)) {
        if ($control -is [System.Windows.Forms.Button]) {
            $wasPrimary = ($control.ForeColor.ToArgb() -eq [System.Drawing.Color]::White.ToArgb() -and $control.BackColor.ToArgb() -ne $defaultButtonBack)
            $wasTinted = ($control.BackColor.ToArgb() -ne $defaultButtonBack -and $control.BackColor.ToArgb() -ne $canvas.ToArgb())
            $isDanger = ([string]$control.Text -match '删除|清空|退出本次')
            $control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $control.FlatAppearance.BorderSize = 1
            $control.UseVisualStyleBackColor = $false
            $control.UseCompatibleTextRendering = $false
            if ($isDanger) {
                $control.BackColor = $dangerSoft
                $control.ForeColor = $danger
                $control.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(232, 176, 180)
                $control.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(253, 223, 225)
                $control.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(248, 207, 210)
            }
            elseif ($wasPrimary) {
                $control.BackColor = $accent
                $control.ForeColor = [System.Drawing.Color]::White
                $control.FlatAppearance.BorderColor = $accent
                $control.FlatAppearance.MouseOverBackColor = $accentHover
                $control.FlatAppearance.MouseDownBackColor = $accentPressed
            }
            elseif ($wasTinted) {
                $control.BackColor = $accentSoft
                $control.ForeColor = $accentPressed
                $control.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(159, 203, 235)
                $control.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(214, 235, 250)
                $control.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(196, 224, 244)
            }
            else {
                $control.BackColor = $surface
                $control.ForeColor = $text
                $control.FlatAppearance.BorderColor = $border
                $control.FlatAppearance.MouseOverBackColor = $surfaceAlt
                $control.FlatAppearance.MouseDownBackColor = $border
            }
        }
        elseif ($control -is [System.Windows.Forms.DataGridView]) {
            $control.EnableHeadersVisualStyles = $false
            $control.BackgroundColor = $surface
            $control.GridColor = $gridLine
            $control.ColumnHeadersDefaultCellStyle.BackColor = $surfaceAlt
            $control.ColumnHeadersDefaultCellStyle.ForeColor = $text
            $control.ColumnHeadersDefaultCellStyle.SelectionBackColor = $surfaceAlt
            $control.ColumnHeadersDefaultCellStyle.SelectionForeColor = $text
            $control.DefaultCellStyle.BackColor = $surface
            $control.DefaultCellStyle.ForeColor = $text
            $control.DefaultCellStyle.SelectionBackColor = $selection
            $control.DefaultCellStyle.SelectionForeColor = $selectionText
            $control.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 254)
            $control.AlternatingRowsDefaultCellStyle.ForeColor = $text
        }
        elseif ($control -is [System.Windows.Forms.TextBox]) {
            $control.BackColor = if ($control.ReadOnly) { $surfaceAlt } else { $surface }
            $control.ForeColor = $text
        }
        elseif ($control -is [System.Windows.Forms.ComboBox]) {
            $control.BackColor = $surface
            $control.ForeColor = $text
            $control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        }
        elseif ($control -is [System.Windows.Forms.ListView] -or
                $control -is [System.Windows.Forms.ListBox] -or
                $control -is [System.Windows.Forms.CheckedListBox]) {
            $control.BackColor = $surface
            $control.ForeColor = $text
        }
        elseif ($control -is [System.Windows.Forms.GroupBox]) {
            $control.BackColor = $canvas
            $control.ForeColor = $text
        }
        elseif ($control -is [System.Windows.Forms.Panel]) {
            $control.BackColor = $canvas
        }
        elseif ($control -is [System.Windows.Forms.Label]) {
            if ($control.BorderStyle -ne [System.Windows.Forms.BorderStyle]::None) {
                $control.BackColor = $accentSoft
                $control.ForeColor = $selectionText
            }
            elseif ($control.ForeColor.ToArgb() -eq [System.Drawing.Color]::DimGray.ToArgb()) {
                $control.ForeColor = $muted
            }
            else {
                $control.ForeColor = $text
            }
        }
        elseif ($control -is [System.Windows.Forms.CheckBox] -or $control -is [System.Windows.Forms.RadioButton]) {
            $control.ForeColor = $text
        }
    }
}

function Update-ExporterProgress {
    param([string]$Message)
    if ($null -ne $script:ExporterProgressCallback -and -not [string]::IsNullOrWhiteSpace($Message)) {
        & $script:ExporterProgressCallback $Message
    }
}

function ConvertTo-SettingBoolean {
    param([object]$Value, [bool]$DefaultValue)
    if ($null -eq $Value) { return $DefaultValue }
    if ($Value -is [bool]) { return [bool]$Value }
    $number = 0
    if ([int]::TryParse([string]$Value, [ref]$number)) { return $number -ne 0 }
    $parsed = $false
    if ([bool]::TryParse([string]$Value, [ref]$parsed)) { return $parsed }
    return $DefaultValue
}

function Set-ObjectPropertyValue {
    param(
        [object]$Object,
        [string]$Name,
        [AllowNull()][object]$Value
    )
    if ($null -eq $Object) { throw '不能向空对象写入工具设置。' }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) { $property.Value = $Value }
    else { $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value }
}

function Get-EmbeddedToolSettings {
    param([string]$ToolPath = $script:ExporterToolPath)
    try {
        if ([string]::IsNullOrWhiteSpace($ToolPath) -or
            [IO.Path]::GetExtension($ToolPath) -ine '.vbs' -or
            -not (Test-Path -LiteralPath $ToolPath -PathType Leaf)) {
            return [pscustomobject]@{}
        }
        $content = [IO.File]::ReadAllText([IO.Path]::GetFullPath($ToolPath), [Text.Encoding]::UTF8)
        $pattern = '(?m)^''#==TOOL_SETTINGS==(?<data>[A-Za-z0-9+/=]*)(?=\r?$)'
        $matches = [regex]::Matches($content, $pattern)
        if ($matches.Count -ne 1 -or [string]::IsNullOrWhiteSpace($matches[0].Groups['data'].Value)) {
            return [pscustomobject]@{}
        }
        $jsonBytes = [Convert]::FromBase64String($matches[0].Groups['data'].Value)
        $json = [Text.Encoding]::UTF8.GetString($jsonBytes)
        $parsed = $json | ConvertFrom-Json
        if ($null -eq $parsed) { return [pscustomobject]@{} }
        return $parsed
    }
    catch { return [pscustomobject]@{} }
}

function Save-EmbeddedToolSettings {
    param(
        [object]$Settings,
        [string]$ToolPath = $script:ExporterToolPath
    )
    $temporaryPath = ''
    try {
        if ([string]::IsNullOrWhiteSpace($ToolPath) -or
            [IO.Path]::GetExtension($ToolPath) -ine '.vbs' -or
            -not (Test-Path -LiteralPath $ToolPath -PathType Leaf)) { return $false }
        $resolvedToolPath = [IO.Path]::GetFullPath($ToolPath)
        $content = [IO.File]::ReadAllText($resolvedToolPath, [Text.Encoding]::UTF8)
        $pattern = '(?m)^''#==TOOL_SETTINGS==(?<data>[A-Za-z0-9+/=]*)(?=\r?$)'
        $matches = [regex]::Matches($content, $pattern)
        if ($matches.Count -ne 1) { return $false }

        $json = $Settings | ConvertTo-Json -Depth 20 -Compress
        if ([string]::IsNullOrWhiteSpace($json)) { $json = '{}' }
        $encoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
        $replacement = $script:ToolSettingsMarker + $encoded
        $match = $matches[0]
        $updated = $content.Substring(0, $match.Index) + $replacement + $content.Substring($match.Index + $match.Length)

        $temporaryPath = Join-Path ([IO.Path]::GetDirectoryName($resolvedToolPath)) ('.tool-settings-' + [guid]::NewGuid().ToString('N') + '.tmp')
        [IO.File]::WriteAllText($temporaryPath, $updated, $script:Utf8NoBom)
        try {
            [IO.File]::Replace($temporaryPath, $resolvedToolPath, $null, $true)
            $temporaryPath = ''
        }
        catch {
            [IO.File]::Copy($temporaryPath, $resolvedToolPath, $true)
            [IO.File]::Delete($temporaryPath)
            $temporaryPath = ''
        }
        $verification = [IO.File]::ReadAllText($resolvedToolPath, [Text.Encoding]::UTF8)
        $verificationMatches = [regex]::Matches($verification, $pattern)
        return $verificationMatches.Count -eq 1 -and $verificationMatches[0].Groups['data'].Value -ceq $encoded
    }
    catch { return $false }
    finally {
        if (-not [string]::IsNullOrWhiteSpace($temporaryPath) -and (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) {
            try { [IO.File]::Delete($temporaryPath) } catch {}
        }
    }
}

function Get-LegacyExporterSettings {
    $key = $null
    try {
        $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($script:LegacyExporterSettingsRegistryPath, $false)
        if ($null -eq $key) { return $null }
        return [pscustomobject][ordered]@{
            Mode = [string]$key.GetValue('Mode', '')
            IncludeCover = $key.GetValue('IncludeCover', $null)
            OpenAfterExport = $key.GetValue('OpenAfterExport', $null)
            AppendFormat = $key.GetValue('AppendFormat', $null)
            AutoNumberDuplicates = $key.GetValue('AutoNumberDuplicates', $null)
            SplitRootGroups = $key.GetValue('SplitRootGroups', $null)
            OutputPath = [string]$key.GetValue('OutputPath', '')
        }
    }
    catch { return $null }
    finally {
        if ($null -ne $key) { $key.Dispose() }
    }
}

function Get-ExporterSettings {
    $settings = [ordered]@{
        Mode = 'Epub'
        IncludeCover = $true
        OpenAfterExport = $true
        AppendFormat = $false
        AutoNumberDuplicates = $false
        SplitRootGroups = $false
        OutputPath = ''
    }
    $rootSettings = Get-EmbeddedToolSettings
    $sourceSettings = Get-ObjectProperty -Object $rootSettings -Name 'exporter' -DefaultValue $null
    $migrateLegacy = $false
    if ($null -eq $sourceSettings) {
        $sourceSettings = Get-LegacyExporterSettings
        $migrateLegacy = $null -ne $sourceSettings
    }
    if ($null -ne $sourceSettings) {
        $savedMode = [string](Get-ObjectProperty -Object $sourceSettings -Name 'Mode' -DefaultValue '')
        if ($savedMode -in @('PerChapter', 'SingleBook', 'Epub', 'PdfStrip')) { $settings.Mode = $savedMode }
        $settings.IncludeCover = ConvertTo-SettingBoolean -Value (Get-ObjectProperty -Object $sourceSettings -Name 'IncludeCover' -DefaultValue $null) -DefaultValue $true
        $settings.OpenAfterExport = ConvertTo-SettingBoolean -Value (Get-ObjectProperty -Object $sourceSettings -Name 'OpenAfterExport' -DefaultValue $null) -DefaultValue $true
        $settings.AppendFormat = ConvertTo-SettingBoolean -Value (Get-ObjectProperty -Object $sourceSettings -Name 'AppendFormat' -DefaultValue $null) -DefaultValue $false
        $settings.AutoNumberDuplicates = ConvertTo-SettingBoolean -Value (Get-ObjectProperty -Object $sourceSettings -Name 'AutoNumberDuplicates' -DefaultValue $null) -DefaultValue $false
        $settings.SplitRootGroups = ConvertTo-SettingBoolean -Value (Get-ObjectProperty -Object $sourceSettings -Name 'SplitRootGroups' -DefaultValue $null) -DefaultValue $false
        $settings.OutputPath = [string](Get-ObjectProperty -Object $sourceSettings -Name 'OutputPath' -DefaultValue '')
    }
    if ($migrateLegacy) {
        [void](Save-ExporterSettings -SavedMode $settings.Mode -IncludeCover $settings.IncludeCover -OpenAfterExport $settings.OpenAfterExport -AppendFormat $settings.AppendFormat -AutoNumberDuplicates $settings.AutoNumberDuplicates -SplitRootGroups $settings.SplitRootGroups -SavedOutputPath $settings.OutputPath)
    }
    return [pscustomobject]$settings
}

function Save-ExporterSettings {
    param(
        [string]$SavedMode,
        [bool]$IncludeCover,
        [bool]$OpenAfterExport,
        [bool]$AppendFormat,
        [bool]$AutoNumberDuplicates,
        [bool]$SplitRootGroups,
        [string]$SavedOutputPath
    )
    try {
        $rootSettings = Get-EmbeddedToolSettings
        $exporterSettings = [pscustomobject][ordered]@{
            Mode = $SavedMode
            IncludeCover = $IncludeCover
            OpenAfterExport = $OpenAfterExport
            AppendFormat = $AppendFormat
            AutoNumberDuplicates = $AutoNumberDuplicates
            SplitRootGroups = $SplitRootGroups
            OutputPath = [string]$SavedOutputPath
        }
        Set-ObjectPropertyValue -Object $rootSettings -Name 'exporter' -Value $exporterSettings
        return Save-EmbeddedToolSettings -Settings $rootSettings
    }
    catch {
        # 设置保存失败不应阻止窗口关闭。
        return $false
    }
}

function Get-ObjectProperty {
    param([object]$Object, [string]$Name, [object]$DefaultValue = $null)
    if ($null -eq $Object) { return $DefaultValue }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $DefaultValue }
    return $property.Value
}

function ConvertTo-MetadataPlainText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    $parts = @($Value | ForEach-Object {
        if ($_ -is [string]) { $_ }
        elseif ($null -ne $_.PSObject.Properties['name']) { [string]$_.name }
        elseif ($null -ne $_.PSObject.Properties['title']) { [string]$_.title }
        else { [string]$_ }
    })
    $text = ($parts -join "`r`n").Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    $text = [regex]::Replace($text, '(?is)<\s*br\s*/?\s*>', "`n")
    $text = [regex]::Replace($text, '(?is)</\s*(?:p|div|li|tr|h[1-6])\s*>', "`n")
    $text = [regex]::Replace($text, '(?is)<\s*li(?:\s[^>]*)?>', '• ')
    $text = [regex]::Replace($text, '(?is)<[^>]+>', '')
    $text = [Net.WebUtility]::HtmlDecode($text)
    $lines = @(($text -replace "`r`n?", "`n") -split "`n" | ForEach-Object { ([regex]::Replace($_, '[\t ]+', ' ')).Trim() })
    return (($lines -join "`r`n").Trim())
}

function Get-MetadataFieldText {
    param(
        [AllowNull()][object]$Metadata,
        [string[]]$Names
    )
    foreach ($name in $Names) {
        $text = ConvertTo-MetadataPlainText -Value (Get-ObjectProperty -Object $Metadata -Name $name -DefaultValue $null)
        if (-not [string]::IsNullOrWhiteSpace($text)) { return $text }
    }
    return ''
}

function Test-ImageFile {
    param([System.IO.FileInfo]$File)
    return $null -ne $File -and $script:ImageExtensions -contains $File.Extension.ToLowerInvariant()
}

function Get-NaturalNameSortKey {
    param([string]$Name)
    return [regex]::Replace([string]$Name, '\d+', {
        param($match)
        return $match.Value.PadLeft(24, '0')
    })
}

function Get-ChapterNameSortKey {
    param([string]$Name)
    $match = [regex]::Match([string]$Name, '(?i)(?:第\s*)?(?<number>\d+(?:\.\d+)?)\s*(?:话|話|章|回|卷|集|ch(?:apter)?)?')
    if ($match.Success) {
        $number = [decimal]0
        if ([decimal]::TryParse($match.Groups['number'].Value, [Globalization.NumberStyles]::AllowDecimalPoint, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
            return '0|' + $number.ToString('00000000000000000000.00000000000000000000', [Globalization.CultureInfo]::InvariantCulture) + '|' + (Get-NaturalNameSortKey $Name)
        }
    }
    return '1|' + (Get-NaturalNameSortKey $Name)
}

function Get-CompositeImageRecord {
    param([System.IO.FileInfo]$File)
    $match = [regex]::Match($File.BaseName, '^(?<prefix>.*\D)(?<page>\d+)$')
    if (-not $match.Success) { return $null }
    $page = [int64]0
    if (-not [int64]::TryParse($match.Groups['page'].Value, [ref]$page)) { return $null }
    return [pscustomobject]@{
        File = $File
        Prefix = $match.Groups['prefix'].Value
        Page = $page
    }
}

function Get-SequenceDisplayName {
    param([string]$Prefix)
    $display = $Prefix.TrimEnd([char[]]' _-.~')
    if ([string]::IsNullOrWhiteSpace($display)) { return $Prefix }
    return $display
}

function Get-FlexibleImageSequence {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$Context
    )
    $issues = New-Object 'System.Collections.Generic.List[string]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    $images = @($Files)
    if ($images.Count -eq 0) {
        $issues.Add(($Context + '：没有图片。'))
        return [pscustomobject]@{ Images = @(); Issues = $issues.ToArray(); Warnings = $warnings.ToArray(); Mode = 'Empty' }
    }
    foreach ($image in $images) {
        if ($image.Length -eq 0) { $issues.Add(($Context + '：图片是空文件：' + $image.Name)) }
    }

    $numericRecords = @()
    $allNumeric = $true
    foreach ($image in $images) {
        if ($image.BaseName -notmatch '^\d+$') { $allNumeric = $false; break }
        $number = [int64]0
        if (-not [int64]::TryParse($image.BaseName, [ref]$number)) {
            $issues.Add(($Context + '：图片编号过大或无效：' + $image.Name))
            continue
        }
        $numericRecords += [pscustomobject]@{ File = $image; Page = $number }
    }
    if ($allNumeric) {
        $duplicates = @($numericRecords | Group-Object Page | Where-Object Count -gt 1)
        foreach ($duplicate in $duplicates) {
            $issues.Add(('{0}：图片编号 {1} 重复（{2}）' -f $Context, $duplicate.Name, (($duplicate.Group.File.Name) -join '、')))
        }
        $orderedNumbers = @($numericRecords.Page | Sort-Object -Unique)
        if ($orderedNumbers.Count -gt 0 -and $orderedNumbers[0] -notin @([int64]0, [int64]1)) {
            $issues.Add(('{0}：第一张图片通常应从 0000 或 0001 开始，实际编号为 {1}。' -f $Context, $orderedNumbers[0]))
        }
        if ($orderedNumbers.Count -gt 0) {
            $span = $orderedNumbers[-1] - $orderedNumbers[0]
            if ($span -le 100000) {
                $numberMap = @{}
                foreach ($number in $orderedNumbers) { $numberMap[[string]$number] = $true }
                for ($number = $orderedNumbers[0]; $number -le $orderedNumbers[-1]; $number++) {
                    if (-not $numberMap.ContainsKey([string]$number)) {
                        $issues.Add(('{0}：缺少图片编号 {1:D4}。' -f $Context, $number))
                    }
                }
            }
            else {
                $issues.Add(($Context + '：图片编号跨度异常，无法逐页检查缺号。'))
            }
        }
        return [pscustomobject]@{
            Images = @($numericRecords | Sort-Object Page, @{ Expression = { $_.File.Name } } | ForEach-Object File)
            Issues = $issues.ToArray()
            Warnings = $warnings.ToArray()
            Mode = 'Numeric'
        }
    }

    $compositeRecords = @()
    foreach ($image in $images) {
        $record = Get-CompositeImageRecord -File $image
        if ($null -eq $record) { $compositeRecords = @(); break }
        $compositeRecords += $record
    }
    if ($compositeRecords.Count -eq $images.Count) {
        $orderedFiles = @()
        $groups = @($compositeRecords | Group-Object Prefix | Sort-Object {
            Get-NaturalNameSortKey -Name (Get-SequenceDisplayName -Prefix $_.Name)
        }, Name)
        foreach ($group in $groups) {
            $groupLabel = Get-SequenceDisplayName -Prefix $group.Name
            $duplicates = @($group.Group | Group-Object Page | Where-Object Count -gt 1)
            foreach ($duplicate in $duplicates) {
                $issues.Add(('{0}：分组 {1} 的页码 {2} 重复（{3}）' -f $Context, $groupLabel, $duplicate.Name, (($duplicate.Group.File.Name) -join '、')))
            }
            $orderedNumbers = @($group.Group.Page | Sort-Object -Unique)
            if ($group.Count -gt 1 -and $orderedNumbers.Count -gt 0 -and $orderedNumbers[0] -notin @([int64]0, [int64]1)) {
                $issues.Add(('{0}：分组 {1} 通常应从 000 或 001 开始，实际从 {2} 开始。' -f $Context, $groupLabel, $orderedNumbers[0]))
            }
            if ($group.Count -gt 1 -and $orderedNumbers.Count -gt 0) {
                $span = $orderedNumbers[-1] - $orderedNumbers[0]
                if ($span -le 100000) {
                    $numberMap = @{}
                    foreach ($number in $orderedNumbers) { $numberMap[[string]$number] = $true }
                    for ($number = $orderedNumbers[0]; $number -le $orderedNumbers[-1]; $number++) {
                        if (-not $numberMap.ContainsKey([string]$number)) {
                            $issues.Add(('{0}：分组 {1} 缺少页码 {2}。' -f $Context, $groupLabel, $number))
                        }
                    }
                }
                else {
                    $issues.Add(('{0}：分组 {1} 的页码跨度异常，无法逐页检查缺号。' -f $Context, $groupLabel))
                }
            }
            $orderedFiles += @($group.Group | Sort-Object Page, @{ Expression = { $_.File.Name } } | ForEach-Object File)
        }
        return [pscustomobject]@{
            Images = @($orderedFiles)
            Issues = $issues.ToArray()
            Warnings = $warnings.ToArray()
            Mode = 'Composite'
        }
    }

    $warnings.Add(($Context + '：无法可靠提取连续页码，已按完整文件名自然排序；所有图片仍会导出，但无法自动证明没有缺图。'))
    return [pscustomobject]@{
        Images = @($images | Sort-Object { Get-NaturalNameSortKey -Name $_.Name }, Name)
        Issues = $issues.ToArray()
        Warnings = $warnings.ToArray()
        Mode = 'NameSorted'
    }
}

function Get-CoverFile {
    param([string]$ComicPath)
    $covers = @(Get-ChildItem -LiteralPath $ComicPath -File -ErrorAction SilentlyContinue | Where-Object {
        $_.BaseName -ieq 'cover' -and (Test-ImageFile -File $_)
    } | Sort-Object {
        switch ($_.Extension.ToLowerInvariant()) {
            '.jpg' { 0 }
            '.jpeg' { 1 }
            '.png' { 2 }
            default { 3 }
        }
    }, Name)
    if ($covers.Count -eq 0) { return $null }
    return $covers[0]
}

function Get-RootBodyImages {
    param([string]$ComicPath)
    return @(Get-ChildItem -LiteralPath $ComicPath -File -ErrorAction SilentlyContinue | Where-Object {
        (Test-ImageFile -File $_) -and $_.BaseName -ine 'cover'
    })
}

function Get-RelativeDirectoryName {
    param([string]$Root, [string]$Path)
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $pathFull = [IO.Path]::GetFullPath($Path)
    if ($pathFull.Length -le $rootFull.Length) { return '' }
    return $pathFull.Substring($rootFull.Length).TrimStart([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
}

function Get-ImageChapterDirectories {
    param([string]$ComicPath)
    $result = New-Object 'System.Collections.Generic.List[object]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    $queue = New-Object 'System.Collections.Generic.Queue[System.IO.DirectoryInfo]'
    foreach ($child in @(Get-ChildItem -LiteralPath $ComicPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ine $script:ReaderResourceFolderName })) { $queue.Enqueue($child) }
    while ($queue.Count -gt 0) {
        $directory = $queue.Dequeue()
        if (($directory.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            $warnings.Add(('已跳过链接目录，避免循环扫描：' + $directory.FullName))
            continue
        }
        $files = @(Get-ChildItem -LiteralPath $directory.FullName -File -ErrorAction SilentlyContinue | Where-Object { Test-ImageFile -File $_ })
        if ($files.Count -gt 0) {
            $relative = Get-RelativeDirectoryName -Root $ComicPath -Path $directory.FullName
            $result.Add([pscustomobject]@{
                Directory = $directory
                RelativeName = $relative
                Files = $files
            })
        }
        foreach ($child in @(Get-ChildItem -LiteralPath $directory.FullName -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ine $script:ReaderResourceFolderName })) { $queue.Enqueue($child) }
    }
    return [pscustomobject]@{ Chapters = $result.ToArray(); Warnings = $warnings.ToArray() }
}

function Get-RootCompositeGroups {
    param([System.IO.FileInfo[]]$Files)
    $records = @()
    $unrecognized = @()
    foreach ($image in @($Files)) {
        $record = Get-CompositeImageRecord -File $image
        if ($null -eq $record) { $unrecognized += $image } else { $records += $record }
    }
    $groups = @($records | Group-Object Prefix)
    if ($records.Count -eq 0 -or $unrecognized.Count -gt 0 -or $groups.Count -lt 2) {
        return [pscustomobject]@{ Recognized = $false; Groups = @() }
    }
    $usedNames = @{}
    $result = @()
    foreach ($group in @($groups | Sort-Object { Get-NaturalNameSortKey -Name (Get-SequenceDisplayName -Prefix $_.Name) }, Name)) {
        $display = Get-SequenceDisplayName -Prefix $group.Name
        if ([string]::IsNullOrWhiteSpace($display)) { $display = '分组' }
        $base = $display
        $suffix = 2
        while ($usedNames.ContainsKey($display)) { $display = $base + ' (' + $suffix + ')'; $suffix++ }
        $usedNames[$display] = $true
        $result += [pscustomobject]@{ Name = $display; Files = @($group.Group.File) }
    }
    return [pscustomobject]@{ Recognized = $true; Groups = @($result) }
}

function Get-WindowsChapterNameMatchKey {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    $text = ([string]$Value).Normalize([Text.NormalizationForm]::FormKC)
    # Windows 与部分下载器会把 ?、: 等非法字符换成全角形式，并移除目录名末尾的空格或句点。
    # 只把这一规则用于元数据匹配，不会改动实际文件夹名称。
    while ($text.Length -gt 0 -and ([char]::IsWhiteSpace($text[$text.Length - 1]) -or $text[$text.Length - 1] -eq '.')) {
        $text = $text.Substring(0, $text.Length - 1)
    }
    return $text
}

function Get-FlexibleChapterLabelPrefixPattern {
    param([string]$Label)
    if ([string]::IsNullOrWhiteSpace($Label)) { return '' }
    $builder = New-Object Text.StringBuilder
    [void]$builder.Append('^')
    foreach ($character in ([regex]::Replace($Label.Trim(), '\s+', '')).ToCharArray()) {
        if ($character -eq '话' -or $character -eq '話') {
            [void]$builder.Append('[话話]')
        }
        else {
            [void]$builder.Append([regex]::Escape([string]$character))
        }
        [void]$builder.Append('\s*')
    }
    return $builder.ToString()
}

function Get-ChapterLabelComparisonKey {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return ([regex]::Replace($Value.Trim(), '\s+', '').Replace('話', '话')).ToLowerInvariant()
}

function Get-ChapterSequenceDisplayText {
    param([string]$Sequence)
    $text = ([string]$Sequence).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    $numericMatch = [regex]::Match($text, '^(\d+(?:\.\d+)?)(?:\s+(.+))?$')
    if (-not $numericMatch.Success) { return $text }
    $suffix = $numericMatch.Groups[2].Value.Trim()
    $label = '第 ' + $numericMatch.Groups[1].Value + ' 话'
    if (-not [string]::IsNullOrWhiteSpace($suffix)) { $label += ' ' + $suffix }
    return $label
}

function Get-CompleteChapterExportLabel {
    param(
        [object]$MetadataItem,
        [string]$Folder,
        [string]$DisplayLabel
    )
    $hasSeparateChapterFields = (
        $null -ne $MetadataItem.PSObject.Properties['chapterSequence'] -and
        $null -ne $MetadataItem.PSObject.Properties['chapterName']
    )
    if ($hasSeparateChapterFields) {
        $sequenceLabel = Get-ChapterSequenceDisplayText -Sequence ([string]$MetadataItem.chapterSequence)
        $chapterName = ([string]$MetadataItem.chapterName).Trim()
        if ([string]::IsNullOrWhiteSpace($sequenceLabel)) { return $chapterName }
        if ([string]::IsNullOrWhiteSpace($chapterName)) { return $sequenceLabel }
        return ($sequenceLabel + ' ' + $chapterName).Trim()
    }

    # 旧元数据没有独立字段时才保留原先的兼容推断。
    $label = ([string]$DisplayLabel).Trim()
    if ([string]::IsNullOrWhiteSpace($label)) { $label = ([string]$Folder).Trim() }

    $titleCandidate = ''
    foreach ($fieldName in @('chapterTitle', 'title', 'chapterName', 'name')) {
        $value = ([string](Get-ObjectProperty -Object $MetadataItem -Name $fieldName -DefaultValue '')).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $titleCandidate = $value
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($titleCandidate)) { $titleCandidate = ([string]$Folder).Trim() }

    $title = $titleCandidate
    $prefixPattern = Get-FlexibleChapterLabelPrefixPattern -Label $label
    if (-not [string]::IsNullOrWhiteSpace($prefixPattern)) {
        $title = [regex]::Replace($titleCandidate, $prefixPattern, '', [Text.RegularExpressions.RegexOptions]::IgnoreCase).Trim()
    }
    if ([string]::IsNullOrWhiteSpace($label)) { return $title }
    if ([string]::IsNullOrWhiteSpace($title)) { return $label }
    $labelKey = Get-ChapterLabelComparisonKey -Value $label
    $titleKey = Get-ChapterLabelComparisonKey -Value $titleCandidate
    if ($labelKey -eq $titleKey -or (-not [string]::IsNullOrWhiteSpace($titleKey) -and $labelKey.EndsWith($titleKey, [StringComparison]::OrdinalIgnoreCase))) {
        return $label
    }
    return ($label + ' ' + $title).Trim()
}

function Read-ComicMetadata {
    param([string]$ComicPath)
    $result = [ordered]@{
        Exists = $false
        StrictOrder = $false
        Infos = @()
        Title = ''
        Author = ''
        Description = ''
        Warning = ''
    }
    $path = Join-Path $ComicPath '元数据.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return [pscustomobject]$result }
    $result.Exists = $true
    try {
        $metadata = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $result.Title = Get-MetadataFieldText -Metadata $metadata -Names @('name', 'title')
        $authorText = Get-MetadataFieldText -Metadata $metadata -Names @('author', 'authors', 'artist', 'artists')
        $result.Author = ([regex]::Replace($authorText, '\s*\r?\n\s*', '、')).Trim('、')
        $result.Description = Get-MetadataFieldText -Metadata $metadata -Names @('description', 'intro', 'summary', 'desc', '简介', '簡介')
        $chapterInfos = @(Get-ObjectProperty -Object $metadata -Name 'chapterInfos' -DefaultValue @())
        if ($chapterInfos.Count -eq 0) {
            $result.Warning = '整理器元数据没有 chapterInfos，已改用名称自然排序。'
            return [pscustomobject]$result
        }
        $infos = @()
        $usedFolders = @{}
        $usedNormalizedFolders = @{}
        $usedOrders = @{}
        foreach ($item in $chapterInfos) {
            $folder = [string](Get-ObjectProperty $item 'chapterFolder' '')
            if ([string]::IsNullOrWhiteSpace($folder)) { $folder = [string](Get-ObjectProperty $item 'chapterTitle' '') }
            $normalizedFolder = Get-WindowsChapterNameMatchKey -Value $folder
            $order = [double]0
            $validOrder = [double]::TryParse([string](Get-ObjectProperty $item 'order' 0), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$order)
            $orderKey = $order.ToString('R', [Globalization.CultureInfo]::InvariantCulture)
            if ([string]::IsNullOrWhiteSpace($folder) -or [string]::IsNullOrWhiteSpace($normalizedFolder) -or -not $validOrder -or $order -le 0 -or $usedFolders.ContainsKey($folder) -or $usedNormalizedFolders.ContainsKey($normalizedFolder) -or $usedOrders.ContainsKey($orderKey)) {
                $result.Warning = '整理器阅读顺序含有空名称、重复名称或重复顺序，已改用名称自然排序。'
                return [pscustomobject]$result
            }
            $label = [string](Get-ObjectProperty $item 'displayLabel' '')
            if ([string]::IsNullOrWhiteSpace($label)) { $label = $folder }
            $completeLabel = Get-CompleteChapterExportLabel -MetadataItem $item -Folder $folder -DisplayLabel $label
            $infos += [pscustomobject]@{ Folder = $folder; MatchKey = $normalizedFolder; Order = $order; Label = $completeLabel; DisplayLabel = $label }
            $usedFolders[$folder] = $true
            $usedNormalizedFolders[$normalizedFolder] = $true
            $usedOrders[$orderKey] = $true
        }
        $result.StrictOrder = $true
        $result.Infos = @($infos | Sort-Object Order)
    }
    catch {
        $result.Warning = '元数据.json 无法解析，已改用名称自然排序：' + $_.Exception.Message
    }
    return [pscustomobject]$result
}

function Get-ComicPlan {
    param(
        [System.IO.DirectoryInfo]$ComicDirectory,
        [switch]$SplitRootGroups
    )
    $issues = New-Object 'System.Collections.Generic.List[string]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    $cover = Get-CoverFile -ComicPath $ComicDirectory.FullName
    $rootImages = @(Get-RootBodyImages -ComicPath $ComicDirectory.FullName)

    $directoryScan = Get-ImageChapterDirectories -ComicPath $ComicDirectory.FullName
    foreach ($message in $directoryScan.Warnings) { $warnings.Add($message) }
    $chapters = @()

    if ($rootImages.Count -gt 0) {
        $rootLayout = Get-RootCompositeGroups -Files $rootImages
        if ($rootLayout.Recognized -and $SplitRootGroups) {
            foreach ($group in $rootLayout.Groups) {
                $sequence = Get-FlexibleImageSequence -Files $group.Files -Context ($ComicDirectory.Name + ' / ' + $group.Name)
                foreach ($message in $sequence.Issues) { $issues.Add($message) }
                foreach ($message in $sequence.Warnings) { $warnings.Add($message) }
                $chapters += [pscustomobject]@{
                    Name = $group.Name
                    RelativeName = $group.Name
                    Label = $group.Name
                    Images = @($sequence.Images)
                    NaturalKey = Get-ChapterNameSortKey $group.Name
                    MetadataOrder = 0
                    MetadataMatched = $false
                    SourceKind = 'RootGroup'
                }
            }
            $warnings.Add(('已按你的导出选项，将根目录图片按文件名前缀拆分为 {0} 个章节组。' -f $rootLayout.Groups.Count))
        }
        else {
            $rootLabel = if (@($directoryScan.Chapters).Count -gt 0) { '根目录正文' } else { '全一话' }
            $sequence = Get-FlexibleImageSequence -Files $rootImages -Context ($ComicDirectory.Name + ' / ' + $rootLabel)
            foreach ($message in $sequence.Issues) { $issues.Add($message) }
            foreach ($message in $sequence.Warnings) { $warnings.Add($message) }
            $chapters += [pscustomobject]@{
                Name = $rootLabel
                RelativeName = $rootLabel
                Label = $rootLabel
                Images = @($sequence.Images)
                NaturalKey = ''
                MetadataOrder = 0
                MetadataMatched = $false
                SourceKind = 'RootBody'
            }
            if ($rootLayout.Recognized) {
                $warnings.Add(('检测到 {0} 个可能的文件名前缀分组；当前按“一个根目录＝一话”导出。' -f $rootLayout.Groups.Count))
            }
        }
    }

    foreach ($chapterSource in @($directoryScan.Chapters)) {
        $sequence = Get-FlexibleImageSequence -Files $chapterSource.Files -Context ($ComicDirectory.Name + ' / ' + $chapterSource.RelativeName)
        foreach ($message in $sequence.Issues) { $issues.Add($message) }
        foreach ($message in $sequence.Warnings) { $warnings.Add($message) }
        $chapters += [pscustomobject]@{
            Name = $chapterSource.Directory.Name
            RelativeName = $chapterSource.RelativeName
            Label = $chapterSource.RelativeName.Replace([IO.Path]::DirectorySeparatorChar, '／')
            Images = @($sequence.Images)
            NaturalKey = Get-ChapterNameSortKey $chapterSource.RelativeName
            MetadataOrder = 0
            MetadataMatched = $false
            SourceKind = 'Directory'
        }
    }

    if ($chapters.Count -eq 0) { $issues.Add('没有找到可导出的图片。') }
    if ($rootImages.Count -gt 0 -and @($directoryScan.Chapters).Count -gt 0) {
        $warnings.Add('根目录正文图片与章节子文件夹同时存在：为避免漏图，两部分都会导出。')
    }

    $metadata = Read-ComicMetadata -ComicPath $ComicDirectory.FullName
    if (-not [string]::IsNullOrWhiteSpace($metadata.Warning)) { $warnings.Add($metadata.Warning) }
    if ($metadata.StrictOrder) {
        $matchedInfo = @{}
        foreach ($chapter in $chapters) {
            $matches = @($metadata.Infos | Where-Object { $_.Folder -ieq $chapter.RelativeName })
            if ($matches.Count -eq 0) { $matches = @($metadata.Infos | Where-Object { $_.Folder -ieq $chapter.Name }) }
            if ($matches.Count -eq 0) {
                $relativeMatchKey = Get-WindowsChapterNameMatchKey -Value $chapter.RelativeName
                $matches = @($metadata.Infos | Where-Object { $_.MatchKey -ceq $relativeMatchKey })
            }
            if ($matches.Count -eq 0) {
                $nameMatchKey = Get-WindowsChapterNameMatchKey -Value $chapter.Name
                $matches = @($metadata.Infos | Where-Object { $_.MatchKey -ceq $nameMatchKey })
            }
            if ($matches.Count -eq 1) {
                $chapter.MetadataOrder = [double]$matches[0].Order
                $chapter.MetadataMatched = $true
                $chapter.Label = [string]$matches[0].Label
                $matchedInfo[[string]$matches[0].Folder] = $true
            }
        }
        $unmatched = @($chapters | Where-Object { -not $_.MetadataMatched })
        if ($unmatched.Count -gt 0) {
            $warnings.Add(('以下章节不在整理器顺序中，已避免漏图并追加到最后：' + (($unmatched | ForEach-Object RelativeName) -join '、')))
        }
        $stale = @($metadata.Infos | Where-Object { -not $matchedInfo.ContainsKey([string]$_.Folder) })
        if ($stale.Count -gt 0) {
            $warnings.Add(('元数据中存在已找不到的章节，已忽略：' + (($stale | ForEach-Object Folder) -join '、')))
        }
        $chapters = @($chapters | Sort-Object @{ Expression = { if ($_.MetadataMatched) { 0 } else { 1 } } }, @{ Expression = { if ($_.MetadataMatched) { $_.MetadataOrder } else { [int]::MaxValue } } }, NaturalKey, RelativeName)
    }
    else {
        $chapters = @($chapters | Sort-Object NaturalKey, RelativeName)
    }

    $total = 0
    foreach ($chapter in $chapters) { $total += @($chapter.Images).Count }
    return [pscustomobject]@{
        Name = $ComicDirectory.Name
        BookTitle = if ([string]::IsNullOrWhiteSpace($metadata.Title)) { $ComicDirectory.Name } else { $metadata.Title }
        Author = $metadata.Author
        Description = $metadata.Description
        SourcePath = $ComicDirectory.FullName
        Cover = $cover
        MetadataUsed = [bool]$metadata.StrictOrder
        Chapters = @($chapters)
        ChapterCount = @($chapters).Count
        TotalImages = $total
        Issues = $issues.ToArray()
        Warnings = $warnings.ToArray()
    }
}

function Test-DirectoryContainsImages {
    param([System.IO.DirectoryInfo]$Directory)
    if ($Directory.Name -ieq $script:ReaderResourceFolderName) { return $false }
    $queue = New-Object 'System.Collections.Generic.Queue[System.IO.DirectoryInfo]'
    $queue.Enqueue($Directory)
    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
        $found = @(Get-ChildItem -LiteralPath $current.FullName -File -ErrorAction SilentlyContinue | Where-Object { Test-ImageFile -File $_ } | Select-Object -First 1)
        if ($found.Count -gt 0) { return $true }
        foreach ($child in @(Get-ChildItem -LiteralPath $current.FullName -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ine $script:ReaderResourceFolderName })) { $queue.Enqueue($child) }
    }
    return $false
}

function Get-CandidateComics {
    param([string]$LibraryRoot, [string]$ResolvedOutputPath)
    $result = @()
    foreach ($directory in @(Get-ChildItem -LiteralPath $LibraryRoot -Directory -ErrorAction Stop | Sort-Object { Get-NaturalNameSortKey $_.Name }, Name)) {
        if (-not [string]::IsNullOrWhiteSpace($ResolvedOutputPath) -and $directory.FullName -ieq $ResolvedOutputPath) { continue }
        if ($directory.Name -in @('.git', 'CBZ导出') -or $directory.Name -ieq $script:ReaderResourceFolderName) { continue }
        if ($directory.Name -ieq 'new' -and (
            (Test-Path -LiteralPath (Join-Path $directory.FullName '漫画整理器.vbs') -PathType Leaf) -or
            (Test-Path -LiteralPath (Join-Path $directory.FullName '漫画整理器.bat') -PathType Leaf)
        )) { continue }
        if (Test-DirectoryContainsImages -Directory $directory) { $result += $directory }
    }
    return @($result)
}

function Get-ShortHash {
    param([string]$Text)
    $sha = [Security.Cryptography.SHA1]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').Substring(0, 8).ToLowerInvariant()
    }
    finally { $sha.Dispose() }
}

function ConvertTo-SafeFileName {
    param([string]$Name, [int]$MaxLength = 100)
    $safe = [regex]::Replace([string]$Name, '[<>:"/\\|?*\x00-\x1F]', '＿').Trim().TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = '未命名' }
    if ($safe -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { $safe = '_' + $safe }
    if ($safe.Length -gt $MaxLength) {
        $hash = Get-ShortHash -Text $safe
        $safe = $safe.Substring(0, [Math]::Max(1, $MaxLength - 10)).TrimEnd() + '~' + $hash
    }
    return $safe
}

function Get-ChapterExportBaseName {
    param(
        [object]$Chapter,
        [int]$Index,
        [int]$Width
    )
    $label = ConvertTo-SafeFileName -Name ([string]$Chapter.Label) -MaxLength 92
    return $Index.ToString(('D' + $Width)) + ' - ' + $label
}

function Test-PathInside {
    param([string]$ChildPath, [string]$ParentPath)
    $child = [IO.Path]::GetFullPath($ChildPath).TrimEnd('\')
    $parent = [IO.Path]::GetFullPath($ParentPath).TrimEnd('\')
    return $child.Equals($parent, [StringComparison]::OrdinalIgnoreCase) -or $child.StartsWith($parent + '\', [StringComparison]::OrdinalIgnoreCase)
}

function Get-ExportTargetDefinition {
    param(
        [string]$ComicName,
        [string]$ResolvedOutputPath,
        [string]$ExportMode,
        [bool]$AppendFormatToFolderName
    )
    $safeComic = ConvertTo-SafeFileName -Name $ComicName -MaxLength 82
    if ($ExportMode -eq 'PdfStrip') {
        $baseName = if ($AppendFormatToFolderName) { $safeComic + ' - PDF' } else { $safeComic }
        return [pscustomobject]@{
            Kind = 'Directory'; BaseName = $baseName; Extension = ''; TargetPath = (Join-Path $ResolvedOutputPath $baseName)
            DetectionBaseNames = @($safeComic, ($safeComic + ' - PDF'))
        }
    }
    if ($ExportMode -eq 'PerChapter') {
        $baseName = if ($AppendFormatToFolderName) { $safeComic + ' - CBZ' } else { $safeComic }
        return [pscustomobject]@{
            Kind = 'Directory'; BaseName = $baseName; Extension = ''; TargetPath = (Join-Path $ResolvedOutputPath $baseName)
            DetectionBaseNames = @($safeComic, ($safeComic + ' - CBZ'))
        }
    }
    $extension = if ($ExportMode -eq 'Epub') { '.epub' } else { '.cbz' }
    return [pscustomobject]@{
        Kind = 'File'; BaseName = $safeComic; Extension = $extension; TargetPath = (Join-Path $ResolvedOutputPath ($safeComic + $extension))
        DetectionBaseNames = @($safeComic)
    }
}

function Find-ExistingExportTarget {
    param([object]$Definition, [string]$ResolvedOutputPath)
    if (-not (Test-Path -LiteralPath $ResolvedOutputPath -PathType Container)) { return '' }
    $escapedNames = @($Definition.DetectionBaseNames | ForEach-Object { [regex]::Escape([string]$_) })
    if ($escapedNames.Count -eq 0) { return '' }
    $numberSuffix = '(?:（[1-9][0-9]*）)?'
    if ($Definition.Kind -eq 'Directory') {
        $pattern = '^(?:' + ($escapedNames -join '|') + ')' + $numberSuffix + '$'
        $match = Get-ChildItem -LiteralPath $ResolvedOutputPath -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $pattern } | Sort-Object Name | Select-Object -First 1
    }
    else {
        $pattern = '^(?:' + ($escapedNames -join '|') + ')' + $numberSuffix + [regex]::Escape([string]$Definition.Extension) + '$'
        $match = Get-ChildItem -LiteralPath $ResolvedOutputPath -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $pattern } | Sort-Object Name | Select-Object -First 1
    }
    if ($null -eq $match) { return '' }
    return [string]$match.FullName
}

function Resolve-ExportTargetPath {
    param(
        [object]$Definition,
        [string]$ResolvedOutputPath,
        [bool]$AllowDuplicateNumbering
    )
    $existing = Find-ExistingExportTarget -Definition $Definition -ResolvedOutputPath $ResolvedOutputPath
    if ([string]::IsNullOrWhiteSpace($existing) -and -not (Test-Path -LiteralPath $Definition.TargetPath)) { return [string]$Definition.TargetPath }
    if (-not $AllowDuplicateNumbering) {
        $shownPath = if ([string]::IsNullOrWhiteSpace($existing)) { [string]$Definition.TargetPath } else { $existing }
        throw ('已经存在同名导出结果：{0}。如需再次导出，请勾选“同名时自动加（1）（2）”。' -f $shownPath)
    }
    for ($number = 1; $number -lt 100000; $number++) {
        $numberedName = [string]$Definition.BaseName + '（' + $number + '）' + [string]$Definition.Extension
        $candidate = Join-Path $ResolvedOutputPath $numberedName
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    throw ('无法为“{0}”找到可用的自动编号。' -f $Definition.BaseName)
}

function New-CbzArchive {
    param(
        [string]$DestinationPath,
        [System.IO.FileInfo[]]$Images,
        [System.IO.FileInfo]$Cover = $null,
        [int]$Digits = 6,
        [string]$ProgressPrefix = '',
        [hashtable]$ImageProgressMap = $null
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $destinationDirectory = [IO.Path]::GetDirectoryName($DestinationPath)
    if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $destinationDirectory -Force)
    }
    $temporaryPath = $DestinationPath + '.tmp-' + [guid]::NewGuid().ToString('N')
    $expected = [ordered]@{}
    try {
        $stream = [IO.File]::Open($temporaryPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $archive = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Create, $false, [Text.Encoding]::UTF8)
            try {
                $page = 0
                if ($null -ne $Cover) {
                    Update-ExporterProgress -Message ($ProgressPrefix + '｜正在写入封面')
                    $coverName = ('{0}-cover{1}' -f ('0' * $Digits), $Cover.Extension.ToLowerInvariant())
                    $entry = $archive.CreateEntry($coverName, [IO.Compression.CompressionLevel]::Optimal)
                    $source = [IO.File]::OpenRead($Cover.FullName)
                    $target = $entry.Open()
                    try { $source.CopyTo($target) } finally { $target.Dispose(); $source.Dispose() }
                    $expected[$coverName] = $Cover.Length
                }
                $bodyImages = @($Images | Where-Object { $null -eq $Cover -or $_.FullName -ine $Cover.FullName })
                for ($imageIndex = 0; $imageIndex -lt $bodyImages.Count; $imageIndex++) {
                    $image = $bodyImages[$imageIndex]
                    $page++
                    $currentPrefix = $ProgressPrefix
                    if ($null -ne $ImageProgressMap -and $ImageProgressMap.ContainsKey($image.FullName)) {
                        $detail = $ImageProgressMap[$image.FullName]
                        $currentPrefix += ('｜第 {0}/{1} 话：{2}｜本话图片 {3}/{4}' -f $detail.ChapterIndex, $detail.ChapterCount, $detail.ChapterLabel, $detail.ImageIndex, $detail.ImageCount)
                    }
                    else {
                        $currentPrefix += ('｜图片 {0}/{1}' -f ($imageIndex + 1), $bodyImages.Count)
                    }
                    Update-ExporterProgress -Message $currentPrefix
                    $entryName = $page.ToString(('D' + $Digits)) + $image.Extension.ToLowerInvariant()
                    $entry = $archive.CreateEntry($entryName, [IO.Compression.CompressionLevel]::Optimal)
                    $source = [IO.File]::OpenRead($image.FullName)
                    $target = $entry.Open()
                    try { $source.CopyTo($target) } finally { $target.Dispose(); $source.Dispose() }
                    $expected[$entryName] = $image.Length
                }
            }
            finally { $archive.Dispose() }
        }
        finally { $stream.Dispose() }

        Update-ExporterProgress -Message ($ProgressPrefix + '｜正在复核 CBZ')
        $check = [IO.Compression.ZipFile]::OpenRead($temporaryPath)
        try {
            $actualEntries = @($check.Entries | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) })
            if ($actualEntries.Count -ne $expected.Count) { throw 'CBZ 复核失败：文件数量不一致。' }
            foreach ($entry in $actualEntries) {
                if (-not $expected.Contains($entry.FullName)) { throw ('CBZ 复核失败：出现意外文件 ' + $entry.FullName) }
                if ([int64]$entry.Length -ne [int64]$expected[$entry.FullName]) { throw ('CBZ 复核失败：文件大小不一致 ' + $entry.FullName) }
                if ($script:ImageExtensions -notcontains [IO.Path]::GetExtension($entry.FullName).ToLowerInvariant()) {
                    throw ('CBZ 复核失败：包含非图片文件 ' + $entry.FullName)
                }
            }
        }
        finally { $check.Dispose() }

        if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) { [IO.File]::Delete($DestinationPath) }
        [IO.File]::Move($temporaryPath, $DestinationPath)
        return [pscustomobject]@{ Path = $DestinationPath; ImageCount = $expected.Count; Bytes = (Get-Item -LiteralPath $DestinationPath).Length }
    }
    catch {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { [IO.File]::Delete($temporaryPath) }
        throw
    }
}

function Get-JpegFrameInfo {
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 10 -or $Bytes[0] -ne 0xFF -or $Bytes[1] -ne 0xD8) { return $null }
    $sofMarkers = @(0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF)
    $position = 2
    while ($position -lt ($Bytes.Length - 1)) {
        while ($position -lt $Bytes.Length -and $Bytes[$position] -ne 0xFF) { $position++ }
        while ($position -lt $Bytes.Length -and $Bytes[$position] -eq 0xFF) { $position++ }
        if ($position -ge $Bytes.Length) { break }
        $marker = [int]$Bytes[$position]
        $position++
        if ($marker -eq 0xD9 -or $marker -eq 0xDA) { break }
        if ($marker -eq 0x01 -or ($marker -ge 0xD0 -and $marker -le 0xD8)) { continue }
        if ($position + 1 -ge $Bytes.Length) { break }
        $segmentLength = (([int]$Bytes[$position]) -shl 8) -bor [int]$Bytes[$position + 1]
        if ($segmentLength -lt 2 -or ($position + $segmentLength) -gt $Bytes.Length) { break }
        if ($sofMarkers -contains $marker -and $segmentLength -ge 8) {
            $height = (([int]$Bytes[$position + 3]) -shl 8) -bor [int]$Bytes[$position + 4]
            $width = (([int]$Bytes[$position + 5]) -shl 8) -bor [int]$Bytes[$position + 6]
            $components = [int]$Bytes[$position + 7]
            if ($width -gt 0 -and $height -gt 0) {
                return [pscustomobject]@{ Width = $width; Height = $height; Components = $components }
            }
            break
        }
        $position += $segmentLength
    }
    return $null
}

function ConvertTo-PdfJpegPayload {
    param([System.IO.FileInfo]$ImageFile)
    $extension = $ImageFile.Extension.ToLowerInvariant()
    if ($extension -in @('.jpg', '.jpeg')) {
        $sourceBytes = [IO.File]::ReadAllBytes($ImageFile.FullName)
        $jpegInfo = Get-JpegFrameInfo -Bytes $sourceBytes
        if ($null -ne $jpegInfo -and $jpegInfo.Components -in @(1, 3)) {
            return [pscustomobject]@{
                Bytes = $sourceBytes
                Width = $jpegInfo.Width
                Height = $jpegInfo.Height
                ColorSpace = if ($jpegInfo.Components -eq 1) { '/DeviceGray' } else { '/DeviceRGB' }
            }
        }
    }

    # PDF 无法直接嵌入 PNG/BMP/GIF 等文件；以高质量 JPEG 重新编码，避免依赖外部软件。
    Add-Type -AssemblyName System.Drawing
    $sourceImage = $null
    $bitmap = $null
    $graphics = $null
    $memory = $null
    $encoderParameters = $null
    try {
        $sourceImage = [Drawing.Image]::FromFile($ImageFile.FullName)
        if ($sourceImage.Width -le 0 -or $sourceImage.Height -le 0) { throw '图片尺寸无效。' }
        $bitmap = New-Object Drawing.Bitmap($sourceImage.Width, $sourceImage.Height, [Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([Drawing.Color]::White)
        $graphics.CompositingQuality = [Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.DrawImage($sourceImage, 0, 0, $sourceImage.Width, $sourceImage.Height)
        $jpegCodec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object MimeType -eq 'image/jpeg' | Select-Object -First 1
        if ($null -eq $jpegCodec) { throw '系统缺少 JPEG 编码器。' }
        $encoderParameters = New-Object Drawing.Imaging.EncoderParameters(1)
        $encoderParameters.Param[0] = New-Object Drawing.Imaging.EncoderParameter([Drawing.Imaging.Encoder]::Quality, [int64]95)
        $memory = New-Object IO.MemoryStream
        $bitmap.Save($memory, $jpegCodec, $encoderParameters)
        return [pscustomobject]@{
            Bytes = $memory.ToArray()
            Width = $sourceImage.Width
            Height = $sourceImage.Height
            ColorSpace = '/DeviceRGB'
        }
    }
    catch {
        throw ('PDF 无法解码图片“{0}”。JPG、PNG、BMP、GIF 可直接处理；WebP/AVIF 需要系统已安装相应图片解码器。原始错误：{1}' -f $ImageFile.FullName, $_.Exception.Message)
    }
    finally {
        if ($null -ne $encoderParameters) { $encoderParameters.Dispose() }
        if ($null -ne $memory) { $memory.Dispose() }
        if ($null -ne $graphics) { $graphics.Dispose() }
        if ($null -ne $bitmap) { $bitmap.Dispose() }
        if ($null -ne $sourceImage) { $sourceImage.Dispose() }
    }
}

function Write-PdfAscii {
    param([IO.FileStream]$Stream, [string]$Text)
    $bytes = [Text.Encoding]::ASCII.GetBytes($Text)
    $Stream.Write($bytes, 0, $bytes.Length)
}

function Get-PdfImagePixelSize {
    param([System.IO.FileInfo]$ImageFile)
    if ($ImageFile.Extension.ToLowerInvariant() -in @('.jpg', '.jpeg')) {
        $jpegInfo = Get-JpegFrameInfo -Bytes ([IO.File]::ReadAllBytes($ImageFile.FullName))
        if ($null -ne $jpegInfo) { return [pscustomobject]@{ Width = $jpegInfo.Width; Height = $jpegInfo.Height } }
    }
    Add-Type -AssemblyName System.Drawing
    $image = $null
    try {
        $image = [Drawing.Image]::FromFile($ImageFile.FullName)
        if ($image.Width -le 0 -or $image.Height -le 0) { throw '图片尺寸无效。' }
        return [pscustomobject]@{ Width = $image.Width; Height = $image.Height }
    }
    catch {
        throw ('PDF 无法读取图片尺寸“{0}”：{1}' -f $ImageFile.FullName, $_.Exception.Message)
    }
    finally {
        if ($null -ne $image) { $image.Dispose() }
    }
}

function New-PdfStripDocument {
    param(
        [string]$DestinationPath,
        [System.IO.FileInfo[]]$Images,
        [System.IO.FileInfo]$Cover = $null,
        [string]$ProgressPrefix = ''
    )
    $destinationDirectory = [IO.Path]::GetDirectoryName($DestinationPath)
    if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $destinationDirectory -Force)
    }
    $orderedImages = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
    if ($null -ne $Cover) { $orderedImages.Add($Cover) }
    foreach ($image in @($Images)) {
        if ($null -ne $Cover -and $image.FullName -ieq $Cover.FullName) { continue }
        $orderedImages.Add($image)
    }
    if ($orderedImages.Count -eq 0) { throw 'PDF 长页没有可写入的图片。' }

    $baseWidth = [double]612
    $logicalHeights = New-Object 'System.Collections.Generic.List[double]'
    $totalHeight = [double]0
    for ($index = 0; $index -lt $orderedImages.Count; $index++) {
        $image = $orderedImages[$index]
        Update-ExporterProgress -Message ($ProgressPrefix + ('｜正在读取图片尺寸 {0}/{1}' -f ($index + 1), $orderedImages.Count))
        $size = Get-PdfImagePixelSize -ImageFile $image
        $height = $baseWidth * ([double]$size.Height / [double]$size.Width)
        $logicalHeights.Add($height)
        $totalHeight += $height
    }
    $scale = if ($totalHeight -gt 14400) { 14400 / $totalHeight } else { [double]1 }
    $pageWidth = $baseWidth * $scale
    $pageHeight = $totalHeight * $scale
    $culture = [Globalization.CultureInfo]::InvariantCulture
    $widthText = $pageWidth.ToString('0.######', $culture)
    $heightText = $pageHeight.ToString('0.######', $culture)

    $temporaryPath = $DestinationPath + '.tmp-' + [guid]::NewGuid().ToString('N')
    $stream = $null
    try {
        $stream = [IO.File]::Open($temporaryPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $contentObject = 4 + $orderedImages.Count
        $objectCount = $contentObject
        $offsets = New-Object 'long[]' ($objectCount + 1)
        Write-PdfAscii -Stream $stream -Text "%PDF-1.4`n%PDFGEN`n"

        $offsets[1] = $stream.Position
        Write-PdfAscii -Stream $stream -Text "1 0 obj`n<< /Type /Catalog /Pages 2 0 R /PageLayout /SinglePage >>`nendobj`n"
        $offsets[2] = $stream.Position
        Write-PdfAscii -Stream $stream -Text "2 0 obj`n<< /Type /Pages /Count 1 /Kids [ 3 0 R ] >>`nendobj`n"

        $xObjects = for ($index = 0; $index -lt $orderedImages.Count; $index++) {
            '/Im' + ($index + 1) + ' ' + (4 + $index) + ' 0 R'
        }
        $offsets[3] = $stream.Position
        Write-PdfAscii -Stream $stream -Text ("3 0 obj`n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {0} {1}] /Resources << /XObject << {2} >> >> /Contents {3} 0 R >>`nendobj`n" -f $widthText, $heightText, ($xObjects -join ' '), $contentObject)

        $contentBuilder = New-Object Text.StringBuilder
        $cursor = $pageHeight
        for ($index = 0; $index -lt $orderedImages.Count; $index++) {
            Update-ExporterProgress -Message ($ProgressPrefix + ('｜正在写入图片 {0}/{1}' -f ($index + 1), $orderedImages.Count))
            $payload = ConvertTo-PdfJpegPayload -ImageFile $orderedImages[$index]
            $imageObject = 4 + $index
            $offsets[$imageObject] = $stream.Position
            Write-PdfAscii -Stream $stream -Text ("{0} 0 obj`n<< /Type /XObject /Subtype /Image /Width {1} /Height {2} /ColorSpace {3} /BitsPerComponent 8 /Filter /DCTDecode /Length {4} >>`nstream`n" -f $imageObject, $payload.Width, $payload.Height, $payload.ColorSpace, $payload.Bytes.Length)
            $stream.Write($payload.Bytes, 0, $payload.Bytes.Length)
            Write-PdfAscii -Stream $stream -Text "`nendstream`nendobj`n"

            $drawHeight = $logicalHeights[$index] * $scale
            $cursor -= $drawHeight
            if ($index -eq ($orderedImages.Count - 1) -or [Math]::Abs($cursor) -lt 0.000001) { $cursor = 0 }
            $drawHeightText = $drawHeight.ToString('0.######', $culture)
            $cursorText = $cursor.ToString('0.######', $culture)
            [void]$contentBuilder.Append("q`n$widthText 0 0 $drawHeightText 0 $cursorText cm`n/Im$($index + 1) Do`nQ`n")
        }

        $contentBytes = [Text.Encoding]::ASCII.GetBytes($contentBuilder.ToString())
        $offsets[$contentObject] = $stream.Position
        Write-PdfAscii -Stream $stream -Text ("{0} 0 obj`n<< /Length {1} >>`nstream`n" -f $contentObject, $contentBytes.Length)
        $stream.Write($contentBytes, 0, $contentBytes.Length)
        Write-PdfAscii -Stream $stream -Text "endstream`nendobj`n"

        $xrefOffset = $stream.Position
        Write-PdfAscii -Stream $stream -Text ("xref`n0 {0}`n0000000000 65535 f `n" -f ($objectCount + 1))
        for ($objectNumber = 1; $objectNumber -le $objectCount; $objectNumber++) {
            Write-PdfAscii -Stream $stream -Text ($offsets[$objectNumber].ToString('D10') + " 00000 n `n")
        }
        Write-PdfAscii -Stream $stream -Text ("trailer`n<< /Size {0} /Root 1 0 R >>`nstartxref`n{1}`n%%EOF`n" -f ($objectCount + 1), $xrefOffset)
        $stream.Dispose()
        $stream = $null

        Update-ExporterProgress -Message ($ProgressPrefix + '｜正在复核 PDF')
        $checkStream = [IO.File]::OpenRead($temporaryPath)
        try {
            if ($checkStream.Length -lt 32) { throw 'PDF 长页复核失败：文件过小。' }
            $headerBytes = New-Object 'byte[]' 8
            if ($checkStream.Read($headerBytes, 0, $headerBytes.Length) -ne $headerBytes.Length -or [Text.Encoding]::ASCII.GetString($headerBytes) -notlike '%PDF-*') {
                throw 'PDF 长页复核失败：文件头无效。'
            }
            $tailLength = [int][Math]::Min(64, $checkStream.Length)
            $tailBytes = New-Object 'byte[]' $tailLength
            [void]$checkStream.Seek(-$tailLength, [IO.SeekOrigin]::End)
            if ($checkStream.Read($tailBytes, 0, $tailBytes.Length) -ne $tailBytes.Length -or [Text.Encoding]::ASCII.GetString($tailBytes) -notmatch '%%EOF') {
                throw 'PDF 长页复核失败：文件没有正常结束。'
            }
        }
        finally { $checkStream.Dispose() }

        if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) { [IO.File]::Delete($DestinationPath) }
        [IO.File]::Move($temporaryPath, $DestinationPath)
        return [pscustomobject]@{ Path = $DestinationPath; ImageCount = $orderedImages.Count; Bytes = (Get-Item -LiteralPath $DestinationPath).Length }
    }
    catch {
        if ($null -ne $stream) { $stream.Dispose() }
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { [IO.File]::Delete($temporaryPath) }
        throw
    }
}

function ConvertTo-XmlText {
    param([object]$Value)
    $text = [regex]::Replace([string]$Value, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return [Security.SecurityElement]::Escape($text)
}

function Get-EpubImageMediaType {
    param([string]$Extension)
    switch ($Extension.ToLowerInvariant()) {
        '.jpg'  { return 'image/jpeg' }
        '.jpeg' { return 'image/jpeg' }
        '.png'  { return 'image/png' }
        '.gif'  { return 'image/gif' }
        '.webp' { return 'image/webp' }
        '.bmp'  { return 'image/bmp' }
        '.avif' { return 'image/avif' }
        default { throw ('EPUB 不支持的图片扩展名：' + $Extension) }
    }
}

function Add-EpubTextEntry {
    param(
        [IO.Compression.ZipArchive]$Archive,
        [string]$EntryName,
        [string]$Text,
        [IO.Compression.CompressionLevel]$Compression = [IO.Compression.CompressionLevel]::Optimal,
        [Text.Encoding]$Encoding = $script:Utf8NoBom
    )
    $entry = $Archive.CreateEntry($EntryName, $Compression)
    $entryStream = $entry.Open()
    try {
        $writer = New-Object IO.StreamWriter($entryStream, $Encoding)
        try { $writer.Write($Text) } finally { $writer.Dispose() }
    }
    finally { $entryStream.Dispose() }
}

function Add-EpubImageEntry {
    param(
        [IO.Compression.ZipArchive]$Archive,
        [string]$EntryName,
        [System.IO.FileInfo]$SourceFile
    )
    $entry = $Archive.CreateEntry($EntryName, [IO.Compression.CompressionLevel]::NoCompression)
    $source = [IO.File]::OpenRead($SourceFile.FullName)
    $target = $entry.Open()
    try { $source.CopyTo($target) } finally { $target.Dispose(); $source.Dispose() }
}

function New-EpubArchive {
    param(
        [string]$DestinationPath,
        [object]$Plan,
        [System.IO.FileInfo]$Cover = $null,
        [string]$ProgressPrefix = ''
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (@($Plan.Chapters).Count -eq 0) { throw 'EPUB 至少需要一个章节。' }

    $destinationDirectory = [IO.Path]::GetDirectoryName($DestinationPath)
    if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $destinationDirectory -Force)
    }
    $temporaryPath = $DestinationPath + '.tmp-' + [guid]::NewGuid().ToString('N')
    $identifier = 'urn:uuid:' + [guid]::NewGuid().ToString()
    $modified = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
    $bookTitle = ConvertTo-XmlText $Plan.BookTitle
    $author = ConvertTo-XmlText $Plan.Author
    $description = ConvertTo-XmlText $Plan.Description
    $manifestItems = @(
        '    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
        '    <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>',
        '    <item id="style" href="styles/book.css" media-type="text/css"/>'
    )
    $spineItems = @()
    $navItems = @()
    $ncxItems = @()
    $expectedImageCount = 0
    $chapterEntryNames = @()

    try {
        $stream = [IO.File]::Open($temporaryPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $archive = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Create, $false, [Text.Encoding]::UTF8)
            try {
                Add-EpubTextEntry -Archive $archive -EntryName 'mimetype' -Text 'application/epub+zip' -Compression ([IO.Compression.CompressionLevel]::NoCompression) -Encoding ([Text.Encoding]::ASCII)

                $containerXml = '<?xml version="1.0" encoding="UTF-8"?>' + "`n" +
                    '<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">' + "`n" +
                    '  <rootfiles><rootfile full-path="EPUB/package.opf" media-type="application/oebps-package+xml"/></rootfiles>' + "`n" +
                    '</container>'
                Add-EpubTextEntry -Archive $archive -EntryName 'META-INF/container.xml' -Text $containerXml

                $css = @'
html, body { margin: 0; padding: 0; background: #000; }
body.comic-page { margin: 0; padding: 0; }
img.comic-image { display: block; width: 100%; height: auto; margin: 0; padding: 0; border: 0; }
body.cover-page { margin: 0; padding: 0; text-align: center; }
img.cover-image { display: block; width: 100%; height: auto; margin: 0 auto; padding: 0; border: 0; }
'@
                Add-EpubTextEntry -Archive $archive -EntryName 'EPUB/styles/book.css' -Text $css

                if ($null -ne $Cover) {
                    Update-ExporterProgress -Message ($ProgressPrefix + '｜正在写入封面')
                    $coverExtension = $Cover.Extension.ToLowerInvariant()
                    $coverEntryName = 'EPUB/images/cover' + $coverExtension
                    Add-EpubImageEntry -Archive $archive -EntryName $coverEntryName -SourceFile $Cover
                    $expectedImageCount++
                    $manifestItems += ('    <item id="cover-image" href="images/cover{0}" media-type="{1}" properties="cover-image"/>' -f $coverExtension, (Get-EpubImageMediaType $coverExtension))
                    $manifestItems += '    <item id="cover-page" href="text/cover.xhtml" media-type="application/xhtml+xml"/>'
                    $spineItems += '    <itemref idref="cover-page" linear="yes"/>'
                    $coverXhtml = '<?xml version="1.0" encoding="UTF-8"?>' + "`n" +
                        '<!DOCTYPE html>' + "`n" +
                        '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="zh-CN" lang="zh-CN">' + "`n" +
                        '<head><meta charset="utf-8"/><title>封面</title><link rel="stylesheet" type="text/css" href="../styles/book.css"/></head>' + "`n" +
                        '<body class="cover-page" epub:type="cover"><img class="cover-image" src="../images/cover' + $coverExtension + '" alt="封面"/></body></html>'
                    Add-EpubTextEntry -Archive $archive -EntryName 'EPUB/text/cover.xhtml' -Text $coverXhtml
                }

                for ($chapterIndex = 0; $chapterIndex -lt $Plan.Chapters.Count; $chapterIndex++) {
                    $chapter = $Plan.Chapters[$chapterIndex]
                    $chapterNumber = $chapterIndex + 1
                    $chapterId = 'chapter-' + $chapterNumber.ToString('D4')
                    $chapterFile = $chapterId + '.xhtml'
                    $chapterEntryName = 'EPUB/text/' + $chapterFile
                    $chapterEntryNames += $chapterEntryName
                    $label = ConvertTo-XmlText $chapter.Label
                    $imageMarkup = @()
                    for ($imageIndex = 0; $imageIndex -lt @($chapter.Images).Count; $imageIndex++) {
                        Update-ExporterProgress -Message ($ProgressPrefix + ('｜第 {0}/{1} 话：{2}｜图片 {3}/{4}' -f $chapterNumber, $Plan.Chapters.Count, $chapter.Label, ($imageIndex + 1), @($chapter.Images).Count))
                        $image = $chapter.Images[$imageIndex]
                        $extension = $image.Extension.ToLowerInvariant()
                        $imageId = 'img-c' + $chapterNumber.ToString('D4') + '-p' + ($imageIndex + 1).ToString('D6')
                        $imageFile = 'c' + $chapterNumber.ToString('D4') + '-p' + ($imageIndex + 1).ToString('D6') + $extension
                        Add-EpubImageEntry -Archive $archive -EntryName ('EPUB/images/' + $imageFile) -SourceFile $image
                        $expectedImageCount++
                        $manifestItems += ('    <item id="{0}" href="images/{1}" media-type="{2}"/>' -f $imageId, $imageFile, (Get-EpubImageMediaType $extension))
                        $imageMarkup += ('<img class="comic-image" src="../images/{0}" alt="{1} - 第 {2} 张"/>' -f $imageFile, $label, ($imageIndex + 1))
                    }
                    $chapterXhtml = '<?xml version="1.0" encoding="UTF-8"?>' + "`n" +
                        '<!DOCTYPE html>' + "`n" +
                        '<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-CN" lang="zh-CN">' + "`n" +
                        '<head><meta charset="utf-8"/><title>' + $label + '</title><link rel="stylesheet" type="text/css" href="../styles/book.css"/></head>' + "`n" +
                        '<body class="comic-page">' + ($imageMarkup -join '') + '</body></html>'
                    Add-EpubTextEntry -Archive $archive -EntryName $chapterEntryName -Text $chapterXhtml
                    $manifestItems += ('    <item id="{0}" href="text/{1}" media-type="application/xhtml+xml"/>' -f $chapterId, $chapterFile)
                    $spineItems += ('    <itemref idref="{0}" linear="yes"/>' -f $chapterId)
                    $navItems += ('      <li><a href="text/{0}">{1}</a></li>' -f $chapterFile, $label)
                    $ncxItems += ('    <navPoint id="nav-{0}" playOrder="{1}"><navLabel><text>{2}</text></navLabel><content src="text/{3}"/></navPoint>' -f $chapterNumber.ToString('D4'), $chapterNumber, $label, $chapterFile)
                }

                $navXhtml = '<?xml version="1.0" encoding="UTF-8"?>' + "`n" +
                    '<!DOCTYPE html>' + "`n" +
                    '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="zh-CN" lang="zh-CN">' + "`n" +
                    '<head><meta charset="utf-8"/><title>目录</title></head><body>' + "`n" +
                    '  <nav epub:type="toc" id="toc"><h1>目录</h1><ol>' + "`n" + ($navItems -join "`n") + "`n" + '    </ol></nav>' + "`n" +
                    '</body></html>'
                Add-EpubTextEntry -Archive $archive -EntryName 'EPUB/nav.xhtml' -Text $navXhtml

                $tocNcx = '<?xml version="1.0" encoding="UTF-8"?>' + "`n" +
                    '<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">' + "`n" +
                    '  <head><meta name="dtb:uid" content="' + $identifier + '"/></head>' + "`n" +
                    '  <docTitle><text>' + $bookTitle + '</text></docTitle><navMap>' + "`n" + ($ncxItems -join "`n") + "`n" + '  </navMap>' + "`n" +
                    '</ncx>'
                Add-EpubTextEntry -Archive $archive -EntryName 'EPUB/toc.ncx' -Text $tocNcx

                $metadataLines = @(
                    '    <dc:identifier id="book-id">' + $identifier + '</dc:identifier>',
                    '    <dc:title>' + $bookTitle + '</dc:title>',
                    '    <dc:language>zh-CN</dc:language>',
                    '    <meta property="dcterms:modified">' + $modified + '</meta>'
                )
                if (-not [string]::IsNullOrWhiteSpace([string]$Plan.Author)) { $metadataLines += '    <dc:creator>' + $author + '</dc:creator>' }
                if (-not [string]::IsNullOrWhiteSpace([string]$Plan.Description)) { $metadataLines += '    <dc:description>' + $description + '</dc:description>' }
                if ($null -ne $Cover) { $metadataLines += '    <meta name="cover" content="cover-image"/>' }
                $guide = if ($null -ne $Cover) { "`n  <guide><reference type=`"cover`" title=`"封面`" href=`"text/cover.xhtml`"/></guide>" } else { '' }
                $packageOpf = '<?xml version="1.0" encoding="UTF-8"?>' + "`n" +
                    '<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="book-id" xml:lang="zh-CN">' + "`n" +
                    '  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">' + "`n" + ($metadataLines -join "`n") + "`n" + '  </metadata>' + "`n" +
                    '  <manifest>' + "`n" + ($manifestItems -join "`n") + "`n" + '  </manifest>' + "`n" +
                    '  <spine toc="ncx">' + "`n" + ($spineItems -join "`n") + "`n" + '  </spine>' + $guide + "`n" +
                    '</package>'
                Add-EpubTextEntry -Archive $archive -EntryName 'EPUB/package.opf' -Text $packageOpf
            }
            finally { $archive.Dispose() }
        }
        finally { $stream.Dispose() }

        Update-ExporterProgress -Message ($ProgressPrefix + '｜正在复核 EPUB')
        $check = [IO.Compression.ZipFile]::OpenRead($temporaryPath)
        try {
            $entries = @($check.Entries)
            if ($entries.Count -eq 0 -or $entries[0].FullName -ne 'mimetype') { throw 'EPUB 复核失败：mimetype 必须是第一个条目。' }
            $required = @('mimetype', 'META-INF/container.xml', 'EPUB/package.opf', 'EPUB/nav.xhtml', 'EPUB/toc.ncx', 'EPUB/styles/book.css') + $chapterEntryNames
            foreach ($entryName in $required) {
                if (@($entries | Where-Object FullName -eq $entryName).Count -ne 1) { throw ('EPUB 复核失败：缺少或重复条目 ' + $entryName) }
            }
            $mimeEntry = $entries | Where-Object FullName -eq 'mimetype' | Select-Object -First 1
            $mimeStream = $mimeEntry.Open()
            try {
                $reader = New-Object IO.StreamReader($mimeStream, [Text.Encoding]::ASCII)
                try { $mimeText = $reader.ReadToEnd() } finally { $reader.Dispose() }
            }
            finally { $mimeStream.Dispose() }
            if ($mimeText -cne 'application/epub+zip') { throw 'EPUB 复核失败：mimetype 内容不正确。' }
            $imageEntries = @($entries | Where-Object { $_.FullName.StartsWith('EPUB/images/') -and -not [string]::IsNullOrWhiteSpace($_.Name) })
            if ($imageEntries.Count -ne $expectedImageCount) { throw 'EPUB 复核失败：图片数量不一致。' }
            foreach ($xmlEntry in @($entries | Where-Object { [IO.Path]::GetExtension($_.FullName).ToLowerInvariant() -in @('.xml', '.opf', '.ncx', '.xhtml') })) {
                $xmlStream = $xmlEntry.Open()
                try {
                    $reader = New-Object IO.StreamReader($xmlStream, [Text.Encoding]::UTF8)
                    try { [void]([xml]$reader.ReadToEnd()) } finally { $reader.Dispose() }
                }
                finally { $xmlStream.Dispose() }
            }
            if (@($entries | Where-Object { $_.FullName -match '(?i)\.json$|(?i)(^|/)index\.html?$' }).Count -gt 0) {
                throw 'EPUB 复核失败：意外包含源 JSON 或 HTML 文件。'
            }
        }
        finally { $check.Dispose() }

        if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) { [IO.File]::Delete($DestinationPath) }
        [IO.File]::Move($temporaryPath, $DestinationPath)
        return [pscustomobject]@{ Path = $DestinationPath; ImageCount = $expectedImageCount; Bytes = (Get-Item -LiteralPath $DestinationPath).Length }
    }
    catch {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { [IO.File]::Delete($temporaryPath) }
        throw
    }
}

function Export-ComicPlan {
    param(
        [object]$Plan,
        [string]$ResolvedOutputPath,
        [string]$ExportMode,
        [bool]$UseCover,
        [bool]$AppendFormatToFolderName,
        [bool]$AllowDuplicateNumbering,
        [int]$ComicIndex = 1,
        [int]$ComicCount = 1
    )
    $comicPrefix = '漫画 {0}/{1}：{2}' -f $ComicIndex, $ComicCount, $Plan.Name
    $targetDefinition = Get-ExportTargetDefinition -ComicName $Plan.Name -ResolvedOutputPath $ResolvedOutputPath -ExportMode $ExportMode -AppendFormatToFolderName $AppendFormatToFolderName
    $resolvedTarget = Resolve-ExportTargetPath -Definition $targetDefinition -ResolvedOutputPath $ResolvedOutputPath -AllowDuplicateNumbering $AllowDuplicateNumbering
    $results = @()
    if ($ExportMode -eq 'PdfStrip') {
        $comicOutput = $resolvedTarget
        $markerPath = Join-Path $comicOutput '.pdf-exporter-owned'
        $managedOutput = Test-Path -LiteralPath $markerPath -PathType Leaf
        if (-not (Test-Path -LiteralPath $comicOutput -PathType Container)) {
            [void](New-Item -ItemType Directory -Path $comicOutput -Force)
            [IO.File]::WriteAllText($markerPath, 'Managed by Local Comic PDF Exporter. Do not place unrelated PDF files in this folder.', $script:Utf8NoBom)
            try { (Get-Item -LiteralPath $markerPath).Attributes = (Get-Item -LiteralPath $markerPath).Attributes -bor [IO.FileAttributes]::Hidden } catch {}
            $managedOutput = $true
        }
        $width = [Math]::Max(3, ([string]$Plan.ChapterCount).Length)
        $usedNames = @{}
        $expectedPaths = @{}
        for ($index = 0; $index -lt $Plan.Chapters.Count; $index++) {
            $chapter = $Plan.Chapters[$index]
            $baseName = Get-ChapterExportBaseName -Chapter $chapter -Index ($index + 1) -Width $width
            $name = $baseName
            $suffix = 2
            while ($usedNames.ContainsKey($name) -or (Test-Path -LiteralPath (Join-Path $comicOutput ($name + '.pdf')) -PathType Container)) {
                $name = $baseName + ' (' + $suffix + ')'
                $suffix++
            }
            $usedNames[$name] = $true
            $target = Join-Path $comicOutput ($name + '.pdf')
            $expectedPaths[$target] = $true
            $cover = if ($UseCover -and $index -eq 0) { $Plan.Cover } else { $null }
            $chapterPrefix = $comicPrefix + ('｜第 {0}/{1} 话：{2}' -f ($index + 1), $Plan.Chapters.Count, $chapter.Label)
            $results += New-PdfStripDocument -DestinationPath $target -Images $chapter.Images -Cover $cover -ProgressPrefix $chapterPrefix
        }
        if ($managedOutput) {
            foreach ($oldFile in @(Get-ChildItem -LiteralPath $comicOutput -File -Filter '*.pdf' -ErrorAction SilentlyContinue)) {
                if (-not $expectedPaths.ContainsKey($oldFile.FullName)) { [IO.File]::Delete($oldFile.FullName) }
            }
        }
        return @($results)
    }
    if ($ExportMode -eq 'Epub') {
        $target = $resolvedTarget
        $cover = if ($UseCover) { $Plan.Cover } else { $null }
        $results += New-EpubArchive -DestinationPath $target -Plan $Plan -Cover $cover -ProgressPrefix $comicPrefix
        return @($results)
    }
    if ($ExportMode -eq 'SingleBook') {
        $allImages = @()
        $imageProgressMap = @{}
        for ($chapterIndex = 0; $chapterIndex -lt $Plan.Chapters.Count; $chapterIndex++) {
            $chapter = $Plan.Chapters[$chapterIndex]
            $chapterImages = @($chapter.Images)
            for ($imageIndex = 0; $imageIndex -lt $chapterImages.Count; $imageIndex++) {
                $image = $chapterImages[$imageIndex]
                $allImages += $image
                $imageProgressMap[$image.FullName] = [pscustomobject]@{
                    ChapterIndex = $chapterIndex + 1
                    ChapterCount = $Plan.Chapters.Count
                    ChapterLabel = $chapter.Label
                    ImageIndex = $imageIndex + 1
                    ImageCount = $chapterImages.Count
                }
            }
        }
        $target = $resolvedTarget
        $cover = if ($UseCover) { $Plan.Cover } else { $null }
        $results += New-CbzArchive -DestinationPath $target -Images $allImages -Cover $cover -Digits 6 -ProgressPrefix $comicPrefix -ImageProgressMap $imageProgressMap
        return @($results)
    }
    $comicOutput = $resolvedTarget
    $markerPath = Join-Path $comicOutput '.cbz-exporter-owned'
    $managedOutput = Test-Path -LiteralPath $markerPath -PathType Leaf
    if (-not (Test-Path -LiteralPath $comicOutput -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $comicOutput -Force)
        [IO.File]::WriteAllText($markerPath, 'Managed by Local Comic CBZ Exporter. Do not place unrelated CBZ files in this folder.', $script:Utf8NoBom)
        try { (Get-Item -LiteralPath $markerPath).Attributes = (Get-Item -LiteralPath $markerPath).Attributes -bor [IO.FileAttributes]::Hidden } catch {}
        $managedOutput = $true
    }
    $width = [Math]::Max(3, ([string]$Plan.ChapterCount).Length)
    $usedNames = @{}
    $expectedPaths = @{}
    for ($index = 0; $index -lt $Plan.Chapters.Count; $index++) {
        $chapter = $Plan.Chapters[$index]
        $baseName = Get-ChapterExportBaseName -Chapter $chapter -Index ($index + 1) -Width $width
        $name = $baseName
        $suffix = 2
        while ($usedNames.ContainsKey($name) -or (Test-Path -LiteralPath (Join-Path $comicOutput ($name + '.cbz')) -PathType Container)) {
            $name = $baseName + ' (' + $suffix + ')'
            $suffix++
        }
        $usedNames[$name] = $true
        $target = Join-Path $comicOutput ($name + '.cbz')
        $expectedPaths[$target] = $true
        $cover = if ($UseCover -and $index -eq 0) { $Plan.Cover } else { $null }
        $chapterPrefix = $comicPrefix + ('｜第 {0}/{1} 话：{2}' -f ($index + 1), $Plan.Chapters.Count, $chapter.Label)
        $results += New-CbzArchive -DestinationPath $target -Images $chapter.Images -Cover $cover -Digits 6 -ProgressPrefix $chapterPrefix
    }
    if ($managedOutput) {
        foreach ($oldFile in @(Get-ChildItem -LiteralPath $comicOutput -File -Filter '*.cbz' -ErrorAction SilentlyContinue)) {
            if (-not $expectedPaths.ContainsKey($oldFile.FullName)) { [IO.File]::Delete($oldFile.FullName) }
        }
    }
    return @($results)
}

function Resolve-OutputPath {
    param([string]$LibraryRoot, [string]$Candidate)
    if ([string]::IsNullOrWhiteSpace($Candidate)) { return [IO.Path]::GetFullPath((Join-Path $LibraryRoot $script:DefaultOutputFolderName)) }
    if ([IO.Path]::IsPathRooted($Candidate)) { return [IO.Path]::GetFullPath($Candidate) }
    return [IO.Path]::GetFullPath((Join-Path $LibraryRoot $Candidate))
}

function Show-LongConfirmation {
    param([string]$Title, [string]$Text)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $dialog = New-Object Windows.Forms.Form
    $dialog.Text = $Title
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object Drawing.Size(780, 560)
    $dialog.MinimizeBox = $false
    $dialog.MaximizeBox = $false
    $box = New-Object Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ReadOnly = $true
    $box.ScrollBars = 'Both'
    $box.WordWrap = $false
    $box.Dock = 'Fill'
    $box.Font = New-Object Drawing.Font('Microsoft YaHei UI', 10)
    $box.Text = $Text
    $panel = New-Object Windows.Forms.FlowLayoutPanel
    $panel.Dock = 'Bottom'
    $panel.Height = 58
    $panel.FlowDirection = 'RightToLeft'
    $panel.Padding = New-Object Windows.Forms.Padding(8)
    $continue = New-Object Windows.Forms.Button
    $continue.Text = '仍然导出'
    $continue.Width = 120
    $continue.Height = 34
    $continue.DialogResult = [Windows.Forms.DialogResult]::Yes
    $cancel = New-Object Windows.Forms.Button
    $cancel.Text = '取消'
    $cancel.Width = 100
    $cancel.Height = 34
    $cancel.DialogResult = [Windows.Forms.DialogResult]::Cancel
    [void]$panel.Controls.Add($continue)
    [void]$panel.Controls.Add($cancel)
    [void]$dialog.Controls.Add($box)
    [void]$dialog.Controls.Add($panel)
    $dialog.AcceptButton = $continue
    $dialog.CancelButton = $cancel
    Set-ComicToolVisualTheme -Window $dialog
    return $dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::Yes
}

function Show-ExporterWindow {
    param([string]$LibraryRoot, [string]$InitialOutput, [switch]$SmokeTest)
    Initialize-ComicToolSharpText
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [Windows.Forms.Application]::EnableVisualStyles()
    $savedSettings = if ($SmokeTest) { [pscustomobject]@{ Mode = 'Epub'; IncludeCover = $true; OpenAfterExport = $true; AppendFormat = $false; AutoNumberDuplicates = $false; SplitRootGroups = $false; OutputPath = '' } } else { Get-ExporterSettings }
    $restoredOutput = if (-not [string]::IsNullOrWhiteSpace($InitialOutput)) { $InitialOutput } elseif (-not [string]::IsNullOrWhiteSpace($savedSettings.OutputPath)) { $savedSettings.OutputPath } else { Resolve-OutputPath -LibraryRoot $LibraryRoot -Candidate '' }

    $form = New-Object Windows.Forms.Form
    $form.Text = '本地漫画 CBZ / EPUB / PDF 导出器'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object Drawing.Size(1040, 875)
    $form.MinimumSize = New-Object Drawing.Size(880, 765)
    $form.Font = New-Object Drawing.Font('Microsoft YaHei UI', 10)

    $header = New-Object Windows.Forms.Label
    $header.Text = '漫画 CBZ / EPUB / PDF 导出器'
    $header.Font = New-Object Drawing.Font('Microsoft YaHei UI', 19, [Drawing.FontStyle]::Bold)
    $header.AutoSize = $true
    $header.Location = New-Object Drawing.Point(24, 18)
    $sub = New-Object Windows.Forms.Label
    $sub.Text = 'JSON 只用于读取章节顺序；源 HTML 不会复制。CBZ 与长页 PDF 均可按话导出。'
    $sub.ForeColor = [Drawing.Color]::DimGray
    $sub.AutoSize = $true
    $sub.Location = New-Object Drawing.Point(27, 60)

    $listLabel = New-Object Windows.Forms.Label
    $listLabel.Text = '选择漫画：'
    $listLabel.AutoSize = $true
    $listLabel.Location = New-Object Drawing.Point(26, 96)
    $list = New-Object Windows.Forms.ListView
    $list.View = [Windows.Forms.View]::Details
    $list.CheckBoxes = $true
    $list.FullRowSelect = $true
    $list.GridLines = $true
    $list.Anchor = 'Top,Bottom,Left,Right'
    $list.Location = New-Object Drawing.Point(28, 124)
    $list.Size = New-Object Drawing.Size(970, 354)
    [void]$list.Columns.Add('漫画文件夹', 650)
    [void]$list.Columns.Add('状态', 290)

    $selectAll = New-Object Windows.Forms.Button
    $selectAll.Text = '全选'
    $selectAll.Location = New-Object Drawing.Point(28, 492)
    $selectAll.Size = New-Object Drawing.Size(90, 34)
    $selectNone = New-Object Windows.Forms.Button
    $selectNone.Text = '全不选'
    $selectNone.Location = New-Object Drawing.Point(126, 492)
    $selectNone.Size = New-Object Drawing.Size(90, 34)
    $refresh = New-Object Windows.Forms.Button
    $refresh.Text = '重新扫描'
    $refresh.Location = New-Object Drawing.Point(224, 492)
    $refresh.Size = New-Object Drawing.Size(110, 34)
    $help = New-Object Windows.Forms.Button
    $help.Text = '使用说明'
    $help.Location = New-Object Drawing.Point(342, 492)
    $help.Size = New-Object Drawing.Size(110, 34)

    $modeGroup = New-Object Windows.Forms.GroupBox
    $modeGroup.Text = '导出方式'
    $modeGroup.Anchor = 'Bottom,Left'
    $modeGroup.Location = New-Object Drawing.Point(28, 528)
    $modeGroup.Size = New-Object Drawing.Size(430, 172)
    $perChapter = New-Object Windows.Forms.RadioButton
    $perChapter.Text = '每话一个 CBZ（推荐）'
    $perChapter.Checked = $savedSettings.Mode -eq 'PerChapter'
    $perChapter.AutoSize = $true
    $perChapter.Location = New-Object Drawing.Point(18, 26)
    $singleBook = New-Object Windows.Forms.RadioButton
    $singleBook.Text = '整部漫画一个 CBZ（无章节目录）'
    $singleBook.Checked = $savedSettings.Mode -eq 'SingleBook'
    $singleBook.AutoSize = $true
    $singleBook.Location = New-Object Drawing.Point(18, 56)
    $epubBook = New-Object Windows.Forms.RadioButton
    $epubBook.Text = '整部 EPUB'
    $epubBook.Checked = $savedSettings.Mode -eq 'Epub'
    $epubBook.AutoSize = $true
    $epubBook.Location = New-Object Drawing.Point(18, 84)
    $pdfStrip = New-Object Windows.Forms.RadioButton
    $pdfStrip.Text = '每话一个 PDF：整话无缝长页（推荐）'
    $pdfStrip.Checked = $savedSettings.Mode -eq 'PdfStrip'
    $pdfStrip.AutoSize = $true
    $pdfStrip.Location = New-Object Drawing.Point(18, 112)
    [void]$modeGroup.Controls.Add($perChapter)
    [void]$modeGroup.Controls.Add($singleBook)
    [void]$modeGroup.Controls.Add($epubBook)
    [void]$modeGroup.Controls.Add($pdfStrip)

    $optionsGroup = New-Object Windows.Forms.GroupBox
    $optionsGroup.Text = '选项'
    $optionsGroup.Anchor = 'Bottom,Left'
    $optionsGroup.Location = New-Object Drawing.Point(470, 528)
    $optionsGroup.Size = New-Object Drawing.Size(390, 172)
    $coverCheck = New-Object Windows.Forms.CheckBox
    $coverCheck.Text = '把总封面放在第一话 / 整本开头'
    $coverCheck.Checked = $savedSettings.IncludeCover
    $coverCheck.AutoSize = $true
    $coverCheck.Location = New-Object Drawing.Point(16, 27)
    $openCheck = New-Object Windows.Forms.CheckBox
    $openCheck.Text = '完成后打开输出文件夹'
    $openCheck.Checked = $savedSettings.OpenAfterExport
    $openCheck.AutoSize = $true
    $openCheck.Location = New-Object Drawing.Point(16, 57)
    $formatFolderCheck = New-Object Windows.Forms.CheckBox
    $formatFolderCheck.Text = '导出文件夹名后标注格式'
    $formatFolderCheck.Checked = $savedSettings.AppendFormat
    $formatFolderCheck.AutoSize = $true
    $formatFolderCheck.Location = New-Object Drawing.Point(16, 87)
    $duplicateCheck = New-Object Windows.Forms.CheckBox
    $duplicateCheck.Text = '同名时自动加（1）（2）继续导出'
    $duplicateCheck.Checked = $savedSettings.AutoNumberDuplicates
    $duplicateCheck.AutoSize = $true
    $duplicateCheck.Location = New-Object Drawing.Point(16, 112)
    $splitRootGroupsCheck = New-Object Windows.Forms.CheckBox
    $splitRootGroupsCheck.Text = '按根目录图片文件名分组拆话'
    $splitRootGroupsCheck.Checked = $savedSettings.SplitRootGroups
    $splitRootGroupsCheck.AutoSize = $true
    $splitRootGroupsCheck.Location = New-Object Drawing.Point(16, 140)
    [void]$optionsGroup.Controls.Add($coverCheck)
    [void]$optionsGroup.Controls.Add($openCheck)
    [void]$optionsGroup.Controls.Add($formatFolderCheck)
    [void]$optionsGroup.Controls.Add($duplicateCheck)
    [void]$optionsGroup.Controls.Add($splitRootGroupsCheck)

    $outputLabel = New-Object Windows.Forms.Label
    $outputLabel.Text = '输出目录：'
    $outputLabel.Anchor = 'Bottom,Left'
    $outputLabel.AutoSize = $true
    $outputLabel.Location = New-Object Drawing.Point(28, 716)
    $outputBox = New-Object Windows.Forms.TextBox
    $outputBox.Anchor = 'Bottom,Left,Right'
    $outputBox.Location = New-Object Drawing.Point(112, 712)
    $outputBox.Size = New-Object Drawing.Size(700, 30)
    $outputBox.Text = $restoredOutput
    $browse = New-Object Windows.Forms.Button
    $browse.Text = '浏览…'
    $browse.Anchor = 'Bottom,Right'
    $browse.Location = New-Object Drawing.Point(822, 710)
    $browse.Size = New-Object Drawing.Size(80, 34)
    $export = New-Object Windows.Forms.Button
    $export.Text = '开始导出'
    $export.Anchor = 'Bottom,Right'
    $export.BackColor = [Drawing.Color]::FromArgb(45, 118, 174)
    $export.ForeColor = [Drawing.Color]::White
    $export.Location = New-Object Drawing.Point(910, 710)
    $export.Size = New-Object Drawing.Size(90, 34)

    $status = New-Object Windows.Forms.Label
    $status.Anchor = 'Bottom,Left,Right'
    $status.AutoSize = $false
    $status.Location = New-Object Drawing.Point(28, 756)
    $status.Size = New-Object Drawing.Size(970, 64)
    $status.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
    $status.Padding = New-Object Windows.Forms.Padding(10, 0, 10, 0)
    $status.TextAlign = [Drawing.ContentAlignment]::MiddleLeft

    foreach ($control in @($header, $sub, $listLabel, $list, $selectAll, $selectNone, $refresh, $help, $modeGroup, $optionsGroup, $outputLabel, $outputBox, $browse, $export, $status)) {
        [void]$form.Controls.Add($control)
    }

    $operationControls = @($list, $selectAll, $selectNone, $refresh, $help, $modeGroup, $optionsGroup, $outputBox, $browse, $export)
    $setExportBusy = {
        param([bool]$Busy)
        $script:ExporterIsRunning = $Busy
        foreach ($control in $operationControls) { $control.Enabled = -not $Busy }
        $form.UseWaitCursor = $Busy
        $form.Cursor = if ($Busy) { [Windows.Forms.Cursors]::WaitCursor } else { [Windows.Forms.Cursors]::Default }
        $form.Refresh()
        [Windows.Forms.Application]::DoEvents()
    }
    $script:ExporterProgressCallback = {
        param([string]$Message)
        if (-not $status.IsDisposed) {
            $status.Text = $Message
            $status.Refresh()
            [Windows.Forms.Application]::DoEvents()
        }
    }

    $script:CandidateItems = @()
    $getCurrentMode = {
        if ($pdfStrip.Checked) { return 'PdfStrip' }
        if ($epubBook.Checked) { return 'Epub' }
        if ($singleBook.Checked) { return 'SingleBook' }
        return 'PerChapter'
    }
    $refreshDuplicateStates = {
        try {
            $resolvedOutput = Resolve-OutputPath -LibraryRoot $LibraryRoot -Candidate $outputBox.Text
            $currentMode = & $getCurrentMode
            $duplicateCount = 0
            foreach ($item in @($list.Items)) {
                if ($null -eq $item.Tag -or $null -eq $item.Tag.Directory) { continue }
                $definition = Get-ExportTargetDefinition -ComicName $item.Tag.Directory.Name -ResolvedOutputPath $resolvedOutput -ExportMode $currentMode -AppendFormatToFolderName $formatFolderCheck.Checked
                $existing = Find-ExistingExportTarget -Definition $definition -ResolvedOutputPath $resolvedOutput
                $isDuplicate = -not [string]::IsNullOrWhiteSpace($existing)
                $item.Tag.IsDuplicate = $isDuplicate
                $item.Tag.ExistingPath = $existing
                if ($isDuplicate) {
                    $duplicateCount++
                    $item.ForeColor = [Drawing.Color]::Gray
                    $item.SubItems[1].Text = '已有同名导出；启用自动编号后可再导出'
                    if (-not $duplicateCheck.Checked -and $item.Checked) { $item.Checked = $false }
                }
                else {
                    $item.ForeColor = [Drawing.SystemColors]::WindowText
                    $item.SubItems[1].Text = if ($item.Tag.HasMetadata) { '可导出；有元数据' } else { '可导出' }
                }
            }
            $status.Text = if ($duplicateCount -gt 0) {
                '找到 {0} 部漫画，其中 {1} 部已有同名导出（灰色）；默认防止重复。' -f $list.Items.Count, $duplicateCount
            }
            else { '找到 {0} 部可导出漫画；未默认勾选，避免误导出。' -f $list.Items.Count }
        }
        catch {
            $status.Text = '无法检查重复项：' + $_.Exception.Message
        }
    }
    $loadCandidates = {
        $form.Cursor = [Windows.Forms.Cursors]::WaitCursor
        $status.Text = '正在扫描漫画文件夹…'
        [Windows.Forms.Application]::DoEvents()
        try {
            $checkedNames = @($list.CheckedItems | ForEach-Object { if ($null -ne $_.Tag) { [string]$_.Tag.Directory.Name } })
            $resolvedOutput = Resolve-OutputPath -LibraryRoot $LibraryRoot -Candidate $outputBox.Text
            $script:CandidateItems = @(Get-CandidateComics -LibraryRoot $LibraryRoot -ResolvedOutputPath $resolvedOutput)
            $list.Items.Clear()
            foreach ($candidate in $script:CandidateItems) {
                $hasMetadata = Test-Path -LiteralPath (Join-Path $candidate.FullName '元数据.json') -PathType Leaf
                $item = New-Object Windows.Forms.ListViewItem($candidate.Name)
                [void]$item.SubItems.Add('可导出')
                $item.Tag = [pscustomobject]@{ Directory = $candidate; HasMetadata = $hasMetadata; IsDuplicate = $false; ExistingPath = '' }
                $item.Checked = $checkedNames -contains $candidate.Name
                [void]$list.Items.Add($item)
            }
            & $refreshDuplicateStates
        }
        finally { $form.Cursor = [Windows.Forms.Cursors]::Default }
    }

    $list.Add_ItemCheck({
        param($sender, $eventArgs)
        $item = $sender.Items[$eventArgs.Index]
        if ($eventArgs.NewValue -eq [Windows.Forms.CheckState]::Checked -and $null -ne $item.Tag -and $item.Tag.IsDuplicate -and -not $duplicateCheck.Checked) {
            $eventArgs.NewValue = $eventArgs.CurrentValue
            $status.Text = '该漫画已有同名导出；如需再次导出，请先勾选“同名时自动加（1）（2）继续导出”。'
        }
    })
    $selectAll.Add_Click({
        foreach ($item in @($list.Items)) {
            if ($null -ne $item.Tag -and (-not $item.Tag.IsDuplicate -or $duplicateCheck.Checked)) { $item.Checked = $true }
        }
    })
    $selectNone.Add_Click({ foreach ($item in @($list.Items)) { $item.Checked = $false } })
    $refresh.Add_Click($loadCandidates)
    $help.Add_Click({
        $helpText = @'
1. “每话一个 CBZ”适合在支持文件夹分组的手机漫画阅读器中阅读；每一话会生成独立 CBZ。
2. “整部漫画一个 CBZ”会把全部图片放进一个文件，但 CBZ 本身没有真正的可点击章节目录。
3. “整部 EPUB”会生成一整本并带章节目录；具体翻页或连续滚动方式由手机阅读软件决定。
4. “每话一个 PDF：整话无缝长页”会为每一话生成一个长页 PDF，图片之间不留空隙。
5. 根目录直接放图片时默认视为一话；只有勾选“按根目录图片文件名分组拆话”才会按前缀拆成多话。
6. 可选择是否把总封面放在开头、完成后打开输出文件夹，以及是否在输出文件夹名后标注格式。
7. 默认会阻止同名重复导出；勾选“同名时自动加（1）（2）”后才会保留多个同名结果。
8. 灰色漫画表示当前输出目录已有同名导出。更换输出目录或启用自动编号后可以再次选择。
9. 导出只读取图片与章节顺序，不会修改源漫画；JSON、HTML 等网页文件不会写入 CBZ、EPUB 或 PDF 正文。
'@
        [Windows.Forms.MessageBox]::Show($form, $helpText, '漫画导出器使用说明', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    })
    foreach ($radio in @($perChapter, $singleBook, $epubBook, $pdfStrip)) {
        $radio.Add_CheckedChanged({ param($sender, $eventArgs) if ($sender.Checked) { & $refreshDuplicateStates } })
    }
    $formatFolderCheck.Add_CheckedChanged({ & $refreshDuplicateStates })
    $duplicateCheck.Add_CheckedChanged({ & $refreshDuplicateStates })
    $outputBox.Add_Leave({ & $refreshDuplicateStates })
    $browse.Add_Click({
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        $dialog.Description = '选择漫画导出目录'
        if (Test-Path -LiteralPath $outputBox.Text -PathType Container) { $dialog.SelectedPath = $outputBox.Text }
        if ($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::OK) { $outputBox.Text = $dialog.SelectedPath; & $refreshDuplicateStates }
        $dialog.Dispose()
    })

    $export.Add_Click({
        $selectedItems = @($list.CheckedItems)
        if ($selectedItems.Count -eq 0) {
            [Windows.Forms.MessageBox]::Show('请至少勾选一部漫画。', '漫画 CBZ / EPUB / PDF 导出器', 'OK', 'Information') | Out-Null
            return
        }
        $blockedDuplicates = @($selectedItems | Where-Object { $null -ne $_.Tag -and $_.Tag.IsDuplicate })
        if ($blockedDuplicates.Count -gt 0 -and -not $duplicateCheck.Checked) {
            [Windows.Forms.MessageBox]::Show('勾选项中存在已经导出的漫画。请取消这些项目，或启用“同名时自动加（1）（2）继续导出”。', '漫画 CBZ / EPUB / PDF 导出器', 'OK', 'Information') | Out-Null
            return
        }
        $resolvedOutput = Resolve-OutputPath -LibraryRoot $LibraryRoot -Candidate $outputBox.Text
        $selectedDirectories = @($selectedItems | ForEach-Object { $_.Tag.Directory })
        foreach ($directory in $selectedDirectories) {
            if (Test-PathInside -ChildPath $resolvedOutput -ParentPath $directory.FullName) {
                [Windows.Forms.MessageBox]::Show(('输出目录不能放在待导出的漫画内部：' + $directory.Name), '漫画 CBZ / EPUB 导出器', 'OK', 'Error') | Out-Null
                return
            }
        }
        & $setExportBusy $true
        try {
            $exportMode = & $getCurrentMode
            if (-not $SmokeTest) {
                Save-ExporterSettings -SavedMode $exportMode -IncludeCover $coverCheck.Checked -OpenAfterExport $openCheck.Checked -AppendFormat $formatFolderCheck.Checked -AutoNumberDuplicates $duplicateCheck.Checked -SplitRootGroups $splitRootGroupsCheck.Checked -SavedOutputPath $outputBox.Text.Trim()
            }
            $plans = @()
            for ($index = 0; $index -lt $selectedDirectories.Count; $index++) {
                $status.Text = ('正在核验 {0}/{1}：{2}' -f ($index + 1), $selectedDirectories.Count, $selectedDirectories[$index].Name)
                [Windows.Forms.Application]::DoEvents()
                $plans += Get-ComicPlan -ComicDirectory $selectedDirectories[$index] -SplitRootGroups:$($splitRootGroupsCheck.Checked)
            }
            $details = New-Object Text.StringBuilder
            foreach ($plan in $plans) {
                if ($plan.Issues.Count -eq 0 -and $plan.Warnings.Count -eq 0) { continue }
                [void]$details.AppendLine(('【{0}】' -f $plan.Name))
                foreach ($message in $plan.Issues) { [void]$details.AppendLine(('  需确认：' + $message)) }
                foreach ($message in $plan.Warnings) { [void]$details.AppendLine(('  提示：' + $message)) }
                [void]$details.AppendLine()
            }
            if ($details.Length -gt 0 -and -not (Show-LongConfirmation -Title '发现需要确认的情况' -Text ($details.ToString() + "`r`n这些问题不会让图片静默丢失。选择【仍然导出】会把已找到的全部图片按显示顺序写入导出文件。"))) {
                $status.Text = '已取消，没有写入文件。'
                return
            }
            [void](New-Item -ItemType Directory -Path $resolvedOutput -Force)
            $filesCreated = 0
            $imagesWritten = 0
            for ($index = 0; $index -lt $plans.Count; $index++) {
                $status.Text = ('正在导出 {0}/{1}：{2}' -f ($index + 1), $plans.Count, $plans[$index].Name)
                [Windows.Forms.Application]::DoEvents()
                $results = @(Export-ComicPlan -Plan $plans[$index] -ResolvedOutputPath $resolvedOutput -ExportMode $exportMode -UseCover $coverCheck.Checked -AppendFormatToFolderName $formatFolderCheck.Checked -AllowDuplicateNumbering $duplicateCheck.Checked -ComicIndex ($index + 1) -ComicCount $plans.Count)
                $filesCreated += $results.Count
                foreach ($result in $results) { $imagesWritten += $result.ImageCount }
            }
            $status.Text = ('完成：{0} 个导出文件，写入 {1} 个图片条目。' -f $filesCreated, $imagesWritten)
            [Windows.Forms.MessageBox]::Show(($status.Text + "`r`n`r`n输出目录：" + $resolvedOutput), '漫画 CBZ / EPUB / PDF 导出器', 'OK', 'Information') | Out-Null
            & $refreshDuplicateStates
            if ($openCheck.Checked) { Start-Process -FilePath 'explorer.exe' -ArgumentList @($resolvedOutput) }
        }
        catch {
            $status.Text = '导出失败：' + $_.Exception.Message
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, '漫画 CBZ / EPUB / PDF 导出器错误', 'OK', 'Error') | Out-Null
        }
        finally {
            & $setExportBusy $false
        }
    })

    $form.Add_FormClosing({
        param($sender, $eventArgs)
        if ($script:ExporterIsRunning) {
            $eventArgs.Cancel = $true
            $status.Text = '导出仍在进行，请等待当前任务完成后再关闭窗口。'
            return
        }
        if (-not $SmokeTest) {
            Save-ExporterSettings -SavedMode (& $getCurrentMode) -IncludeCover $coverCheck.Checked -OpenAfterExport $openCheck.Checked -AppendFormat $formatFolderCheck.Checked -AutoNumberDuplicates $duplicateCheck.Checked -SplitRootGroups $splitRootGroupsCheck.Checked -SavedOutputPath $outputBox.Text.Trim()
        }
    })
    & $loadCandidates
    Set-ComicToolVisualTheme -Window $form
    if ($SmokeTest) {
        if ($help.Text -ne '使用说明') { throw '导出器缺少“使用说明”按钮。' }
        $timer = New-Object Windows.Forms.Timer
        $timer.Interval = 350
        $timer.Add_Tick({ $timer.Stop(); $form.Close() })
        $form.Add_Shown({ $timer.Start() })
    }
    try { [void]$form.ShowDialog() }
    finally {
        $script:ExporterProgressCallback = $null
        $script:ExporterIsRunning = $false
    }
}

try {
    $toolPath = ''
    if (-not [string]::IsNullOrWhiteSpace($env:LOCAL_COMIC_TOOL_PATH)) { $toolPath = $env:LOCAL_COMIC_TOOL_PATH }
    elseif ($null -ne $MyInvocation.MyCommand.PSObject.Properties['Path']) { $toolPath = [string]$MyInvocation.MyCommand.Path }
    $script:ExporterToolPath = $toolPath
    $scriptDirectory = if ([string]::IsNullOrWhiteSpace($toolPath)) { (Get-Location).Path } else { [IO.Path]::GetDirectoryName($toolPath) }
    $rootCandidate = if ([string]::IsNullOrWhiteSpace($RootPath)) { $scriptDirectory } elseif ([IO.Path]::IsPathRooted($RootPath)) { $RootPath } else { Join-Path $scriptDirectory $RootPath }
    $resolvedRoot = (Resolve-Path -LiteralPath $rootCandidate).Path
    $resolvedOutput = Resolve-OutputPath -LibraryRoot $resolvedRoot -Candidate $OutputPath

    if ($NonInteractive) {
        if ([string]::IsNullOrWhiteSpace($ComicName)) { throw '非交互模式必须提供 -ComicName。' }
        $comicPath = Join-Path $resolvedRoot $ComicName
        if (-not (Test-Path -LiteralPath $comicPath -PathType Container)) { throw ('漫画文件夹不存在：' + $comicPath) }
        if (Test-PathInside -ChildPath $resolvedOutput -ParentPath $comicPath) { throw '输出目录不能放在待导出的漫画内部。' }
        $plan = Get-ComicPlan -ComicDirectory (Get-Item -LiteralPath $comicPath) -SplitRootGroups:$SplitRootGroups
        foreach ($message in $plan.Issues) { Write-Host ('[需确认] ' + $message) -ForegroundColor Yellow }
        foreach ($message in $plan.Warnings) { Write-Host ('[提示] ' + $message) -ForegroundColor DarkYellow }
        Write-Host ('[核验] {0}：{1} 话，{2} 张正文图片，元数据顺序={3}' -f $plan.Name, $plan.ChapterCount, $plan.TotalImages, $plan.MetadataUsed)
        if ($ValidateOnly) { exit 0 }
        if ($plan.Issues.Count -gt 0 -and -not $ForceIssues) { throw '存在需要确认的编号或文件问题；非交互导出请添加 -ForceIssues。' }
        [void](New-Item -ItemType Directory -Path $resolvedOutput -Force)
        $results = @(Export-ComicPlan -Plan $plan -ResolvedOutputPath $resolvedOutput -ExportMode $Mode -UseCover ([bool]$IncludeCover) -AppendFormatToFolderName ([bool]$AppendFormatToFolderName) -AllowDuplicateNumbering ([bool]$AutoNumberDuplicates))
        foreach ($result in $results) { Write-Host ('[完成] {0}（{1} 个图片条目）' -f $result.Path, $result.ImageCount) -ForegroundColor Green }
        exit 0
    }

    $guiInitialOutput = if ([string]::IsNullOrWhiteSpace($OutputPath)) { '' } else { $resolvedOutput }
    Show-ExporterWindow -LibraryRoot $resolvedRoot -InitialOutput $guiInitialOutput -SmokeTest:$UiSmokeTest
    exit 0
}
catch {
    Write-Host ('[错误] ' + $_.Exception.Message) -ForegroundColor Red
    if ($NonInteractive) {
        if (-not [string]::IsNullOrWhiteSpace($_.InvocationInfo.PositionMessage)) { Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor DarkGray }
        if (-not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) { Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray }
    }
    if (-not $NonInteractive -and -not $UiSmokeTest) {
        Add-Type -AssemblyName System.Windows.Forms
        [Windows.Forms.MessageBox]::Show($_.Exception.Message, '漫画 CBZ / EPUB / PDF 导出器错误', 'OK', 'Error') | Out-Null
    }
    exit 1
}
