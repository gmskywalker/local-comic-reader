[CmdletBinding()]
param(
    [string]$RootPath = '',
    [string]$PlanPath = '',
    [switch]$ValidateOnly,
    [switch]$NonInteractive,
    [switch]$UiSmokeTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ImageExtensions = @('.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.avif')
$script:OutputFolderName = '整理完成'
$script:RootChapterToken = '[根目录正文]'
$script:ReaderResourceFolderName = '漫画阅读器资源'
$script:ChapterCoverFolderName = '章节封面'
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:OrganizerProgressCallback = $null
$script:OrganizerIsRunning = $false
$script:OrganizerBusyControlStates = @()

function Update-OrganizerProgress {
    param([string]$Message)
    if ($null -ne $script:OrganizerProgressCallback -and -not [string]::IsNullOrWhiteSpace($Message)) {
        & $script:OrganizerProgressCallback $Message
    }
}

function Write-JsonFile {
    param([string]$Path, [object]$Value)
    $json = $Value | ConvertTo-Json -Depth 20
    [IO.File]::WriteAllText($Path, $json, $script:Utf8NoBom)
}

function Get-ObjectProperty {
    param(
        [object]$Object,
        [string]$Name,
        [AllowNull()][object]$Default = $null
    )
    if ($null -ne $Object -and $null -ne $Object.PSObject.Properties[$Name]) {
        return $Object.$Name
    }
    return $Default
}

function Test-SimpleFolderName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name -in @('.', '..')) { return $false }
    if ($Name -match '[<>:"/\\|?*]') { return $false }
    if ($Name.EndsWith(' ') -or $Name.EndsWith('.')) { return $false }
    return ([IO.Path]::GetFileName($Name) -ceq $Name)
}

function ConvertTo-ChapterNumberInfo {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ($text -notmatch '^\d+(?:\.\d+)?$') { return $null }
    $number = [decimal]0
    $parsed = [decimal]::TryParse(
        $text,
        [Globalization.NumberStyles]::AllowDecimalPoint,
        [Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )
    if (-not $parsed -or $number -le 0) { return $null }
    return [pscustomobject]@{
        Text = $number.ToString('0.############################', [Globalization.CultureInfo]::InvariantCulture)
        Value = $number
    }
}

function Get-ChapterNumberFromName {
    param([string]$Name)
    $match = [regex]::Match($Name, '^第\s*(\d+(?:\.\d+)?)\s*[话話]')
    if (-not $match.Success) { return $null }
    return ConvertTo-ChapterNumberInfo -Value $match.Groups[1].Value
}

function Get-ChapterTitleFromName {
    param([string]$Name)
    $match = [regex]::Match($Name, '^第\s*\d+(?:\.\d+)?\s*[话話]\s*(.*)$')
    if (-not $match.Success) { return '' }
    return $match.Groups[1].Value.Trim()
}

function Get-SpecialChapterInfoFromName {
    param([string]$Name)
    $match = [regex]::Match(
        $Name,
        '^(特典(?:话|話)?|番外(?:篇|话|話)?(?:\s*\d+(?:\.\d+)?)?|插画集|插畫集|イラスト集|附录|附錄|附赠|附贈|后日谈|後日談|后记|後記|小短篇|封面|线下画展展品|線下畫展展品|Extra|Special)(?:\s+(.*))?$',
        [Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    if (-not $match.Success) { return $null }
    return [pscustomobject]@{
        Label = $match.Groups[1].Value.Trim()
        Title = $match.Groups[2].Value.Trim()
    }
}

function ConvertTo-ChapterLabelInfo {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    $numberInfo = ConvertTo-ChapterNumberInfo -Value $text
    if ($null -ne $numberInfo) {
        return [pscustomobject]@{
            IsNumeric = $true
            Number = $numberInfo.Text
            SortNumber = $numberInfo.Value
            BaseFolderName = ('第' + $numberInfo.Text + '话')
            DisplayLabel = ('第 ' + $numberInfo.Text + ' 话')
        }
    }
    if (-not (Test-SimpleFolderName -Name $text) -or $text -match '^\d') {
        return $null
    }
    return [pscustomobject]@{
        IsNumeric = $false
        Number = ''
        SortNumber = [decimal]0
        BaseFolderName = $text
        DisplayLabel = $text
    }
}

function Get-InitialOrganizerChapterFields {
    param(
        [string]$ChapterName,
        [bool]$IsRootChapter,
        [string]$DefaultNumber,
        [bool]$PreserveNumericNumber
    )
    $initialNumber = $DefaultNumber
    $sourceNumber = Get-ChapterNumberFromName -Name $ChapterName
    $specialInfo = Get-SpecialChapterInfoFromName -Name $ChapterName
    if ($null -ne $sourceNumber -and ($PreserveNumericNumber -or $sourceNumber.Text.Contains('.'))) {
        $initialNumber = $sourceNumber.Text
    }
    elseif ($null -ne $specialInfo) {
        $initialNumber = $specialInfo.Label
    }
    elseif ($null -eq $sourceNumber -and $ChapterName -cne $script:RootChapterToken) {
        $arbitraryLabel = ConvertTo-ChapterLabelInfo -Value $ChapterName
        if ($null -ne $arbitraryLabel) { $initialNumber = $ChapterName }
    }
    $initialTitle = if ($null -ne $specialInfo) { $specialInfo.Title } else { Get-ChapterTitleFromName -Name $ChapterName }
    return [pscustomobject]@{ Number = $initialNumber; Title = $initialTitle }
}

function ConvertTo-NumberRangeText {
    param(
        [int[]]$Numbers,
        [int]$MaximumRanges = 12
    )
    $ordered = @($Numbers | Sort-Object -Unique)
    if ($ordered.Count -eq 0) { return '' }
    $ranges = New-Object 'System.Collections.Generic.List[string]'
    $rangeStart = $ordered[0]
    $previous = $ordered[0]
    for ($index = 1; $index -le $ordered.Count; $index++) {
        $current = if ($index -lt $ordered.Count) { $ordered[$index] } else { $null }
        if ($null -ne $current -and $current -eq ($previous + 1)) {
            $previous = $current
            continue
        }
        if ($rangeStart -eq $previous) {
            $ranges.Add($rangeStart.ToString('D4'))
        }
        else {
            $ranges.Add(($rangeStart.ToString('D4') + '-' + $previous.ToString('D4')))
        }
        if ($null -ne $current) {
            $rangeStart = $current
            $previous = $current
        }
    }
    if ($ranges.Count -le $MaximumRanges) {
        return ($ranges -join '、')
    }
    return ((@($ranges | Select-Object -First $MaximumRanges) -join '、') + '、……')
}

function Get-NaturalNameSortKey {
    param([string]$Name)
    return [regex]::Replace($Name, '\d+', {
        param($match)
        return $match.Value.PadLeft(24, '0')
    })
}

function Get-WindowsChapterNameMatchKey {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    $text = ([string]$Value).Normalize([Text.NormalizationForm]::FormKC)
    while ($text.Length -gt 0 -and ([char]::IsWhiteSpace($text[$text.Length - 1]) -or $text[$text.Length - 1] -eq '.')) {
        $text = $text.Substring(0, $text.Length - 1)
    }
    return $text
}

function Get-SourceChapterOrderConfiguration {
    param([string]$ComicPath)
    $result = [ordered]@{ Enabled = $false; ExactMap = @{}; NormalizedMap = @{}; Warning = ''; ShowChapterCovers = $false }
    $metadataPath = Join-Path $ComicPath '元数据.json'
    if (-not (Test-Path -LiteralPath $metadataPath -PathType Leaf)) { return [pscustomobject]$result }
    try {
        $metadata = [IO.File]::ReadAllText($metadataPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($null -ne $metadata.PSObject.Properties['readerOptions'] -and $null -ne $metadata.readerOptions -and $null -ne $metadata.readerOptions.PSObject.Properties['showChapterCovers']) {
            $result.ShowChapterCovers = [bool]$metadata.readerOptions.showChapterCovers
        }
        $chapterInfos = if ($null -ne $metadata.PSObject.Properties['chapterInfos']) { @($metadata.chapterInfos) } else { @() }
        if ($chapterInfos.Count -eq 0) { return [pscustomobject]$result }
        $errors = New-Object 'System.Collections.Generic.List[string]'
        $exactMap = @{}
        $normalizedMap = @{}
        $usedOrders = @{}
        foreach ($chapterInfo in $chapterInfos) {
            $chapterFolder = ''
            if ($null -ne $chapterInfo.PSObject.Properties['chapterFolder']) { $chapterFolder = [string]$chapterInfo.chapterFolder }
            elseif ($null -ne $chapterInfo.PSObject.Properties['chapterTitle']) { $chapterFolder = [string]$chapterInfo.chapterTitle }
            $order = 0
            if ([string]::IsNullOrWhiteSpace($chapterFolder) -or $null -eq $chapterInfo.PSObject.Properties['order'] -or -not [int]::TryParse([string]$chapterInfo.order, [ref]$order) -or $order -lt 1) {
                $errors.Add('存在缺少章节文件夹名或有效顺序的 chapterInfos 项。')
                continue
            }
            if ($exactMap.ContainsKey($chapterFolder)) {
                $errors.Add(('章节名称重复：' + $chapterFolder))
                continue
            }
            if ($usedOrders.ContainsKey([string]$order)) {
                $errors.Add(('章节顺序重复：' + $order))
                continue
            }
            $normalizedKey = Get-WindowsChapterNameMatchKey -Value $chapterFolder
            if ([string]::IsNullOrWhiteSpace($normalizedKey)) {
                $errors.Add(('章节名称无法用于匹配：' + $chapterFolder))
                continue
            }
            if ($normalizedMap.ContainsKey($normalizedKey)) {
                $errors.Add(('章节名称按 Windows 文件名规则处理后重复：{0}、{1}' -f $normalizedMap[$normalizedKey].Folder, $chapterFolder))
                continue
            }
            $coverMode = if ($null -ne $chapterInfo.PSObject.Properties['coverMode']) { ([string]$chapterInfo.coverMode).Trim().ToLowerInvariant() } else { 'first' }
            if ($coverMode -notin @('first', 'custom', 'none')) { $coverMode = 'first' }
            $coverFile = if ($null -ne $chapterInfo.PSObject.Properties['coverFile']) { ([string]$chapterInfo.coverFile).Trim() } else { '' }
            $entry = [pscustomobject]@{ Folder = $chapterFolder; Order = $order; CoverMode = $coverMode; CoverFile = $coverFile }
            $exactMap[$chapterFolder] = $entry
            $normalizedMap[$normalizedKey] = $entry
            $usedOrders[[string]$order] = $true
        }
        if ($errors.Count -gt 0) {
            $result.Warning = '元数据.json 的章节顺序无效，已改用名称自然排序：' + ($errors -join '；')
            return [pscustomobject]$result
        }
        $result.Enabled = $true
        $result.ExactMap = $exactMap
        $result.NormalizedMap = $normalizedMap
    }
    catch {
        $result.Warning = '元数据.json 无法读取章节顺序，已改用名称自然排序：' + $_.Exception.Message
    }
    return [pscustomobject]$result
}

function Set-SourceChapterEntryOrder {
    param([string]$ComicPath, [object[]]$Entries)
    $naturalEntries = @($Entries | Sort-Object { Get-NaturalNameSortKey -Name $_.Name }, Name)
    $configuration = Get-SourceChapterOrderConfiguration -ComicPath $ComicPath
    if (-not $configuration.Enabled) {
        foreach ($entry in $naturalEntries) {
            $entry | Add-Member -NotePropertyName OrderSource -NotePropertyValue 'Natural' -Force
            $entry | Add-Member -NotePropertyName OrderWarning -NotePropertyValue $configuration.Warning -Force
        }
        return @($naturalEntries)
    }
    $matched = @()
    $missing = @()
    $usedMetadataFolders = @{}
    foreach ($entry in $naturalEntries) {
        $configured = $null
        if ($configuration.ExactMap.ContainsKey($entry.Name)) {
            $configured = $configuration.ExactMap[$entry.Name]
        }
        else {
            $normalizedKey = Get-WindowsChapterNameMatchKey -Value $entry.Name
            if ($configuration.NormalizedMap.ContainsKey($normalizedKey)) { $configured = $configuration.NormalizedMap[$normalizedKey] }
        }
        if ($null -eq $configured) {
            $missing += $entry.Name
            continue
        }
        $usedMetadataFolders[[string]$configured.Folder] = $true
        $entry | Add-Member -NotePropertyName ConfiguredCoverMode -NotePropertyValue ([string]$configured.CoverMode) -Force
        $entry | Add-Member -NotePropertyName ConfiguredCoverFile -NotePropertyValue ([string]$configured.CoverFile) -Force
        $entry | Add-Member -NotePropertyName ConfiguredShowChapterCovers -NotePropertyValue ([bool]$configuration.ShowChapterCovers) -Force
        $matched += [pscustomobject]@{ Entry = $entry; Order = [int]$configured.Order }
    }
    if ($missing.Count -gt 0) {
        $warning = '元数据.json 的章节顺序没有覆盖以下章节，已改用名称自然排序：' + ($missing -join '、')
        foreach ($entry in $naturalEntries) {
            $entry | Add-Member -NotePropertyName OrderSource -NotePropertyValue 'Natural' -Force
            $entry | Add-Member -NotePropertyName OrderWarning -NotePropertyValue $warning -Force
        }
        return @($naturalEntries)
    }
    $stale = @($configuration.ExactMap.Keys | Where-Object { -not $usedMetadataFolders.ContainsKey($_) })
    $warning = if ($stale.Count -gt 0) { '元数据顺序中有已不存在的章节，已忽略：' + ($stale -join '、') } else { '' }
    $ordered = @($matched | Sort-Object Order | ForEach-Object { $_.Entry })
    foreach ($entry in $ordered) {
        $entry | Add-Member -NotePropertyName OrderSource -NotePropertyValue 'Metadata' -Force
        $entry | Add-Member -NotePropertyName OrderWarning -NotePropertyValue $warning -Force
    }
    return @($ordered)
}

function Get-NumericImages {
    param(
        [string]$DirectoryPath,
        [switch]$IgnoreCover
    )
    $errors = New-Object 'System.Collections.Generic.List[string]'
    $images = @(Get-ChildItem -LiteralPath $DirectoryPath -File | Where-Object {
        $script:ImageExtensions -contains $_.Extension.ToLowerInvariant() -and
        (-not $IgnoreCover -or $_.BaseName -ine 'cover')
    })
    foreach ($image in $images) {
        if ($image.BaseName -notmatch '^\d+$') {
            $errors.Add(('图片名称不是纯数字：' + $image.Name))
        }
    }
    if ($errors.Count -gt 0) {
        throw ($errors -join '；')
    }
    $ordered = @($images | Sort-Object { [int64]$_.BaseName }, Name)
    if ($ordered.Count -eq 0) {
        throw '没有图片。'
    }
    $numberMap = @{}
    foreach ($image in $ordered) {
        $number = [int64]$image.BaseName
        if ($numberMap.ContainsKey([string]$number)) {
            throw ('图片数字编号重复：' + $number)
        }
        $numberMap[[string]$number] = $image
    }
    $firstNumber = [int64]$ordered[0].BaseName
    $lastNumber = [int64]$ordered[-1].BaseName
    if ($firstNumber -ne 1) {
        throw ('图片编号应从 0001 开始，实际从 ' + $ordered[0].Name + ' 开始。')
    }
    for ($number = 1; $number -le $lastNumber; $number++) {
        if (-not $numberMap.ContainsKey([string]$number)) {
            throw ('缺少图片编号：' + $number.ToString('D4'))
        }
    }
    return @($ordered)
}

function Get-RootBodyImageFiles {
    param([string]$ComicPath)
    return @(Get-ChildItem -LiteralPath $ComicPath -File -ErrorAction SilentlyContinue | Where-Object {
        $script:ImageExtensions -contains $_.Extension.ToLowerInvariant() -and
        $_.BaseName -ine 'cover'
    })
}

function Get-CompositeImageRecord {
    param([System.IO.FileInfo]$File)
    $match = [regex]::Match($File.BaseName, '^(?<prefix>.*\D)(?<page>\d+)$')
    if (-not $match.Success) { return $null }
    $page = [int64]0
    if (-not [int64]::TryParse($match.Groups['page'].Value, [ref]$page)) { return $null }
    return [pscustomobject]@{ File = $File; Prefix = $match.Groups['prefix'].Value; Page = $page }
}

function Get-SequenceDisplayName {
    param([string]$Prefix)
    $display = $Prefix.TrimEnd([char[]]' _-.~')
    if ([string]::IsNullOrWhiteSpace($display)) { $display = '分组' }
    $display = ($display -replace '[<>:"/\\|?*]', '_').TrimEnd([char[]]' .')
    if ([string]::IsNullOrWhiteSpace($display)) { return '分组' }
    return $display
}

function Get-FlexibleImageSequence {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$Context
    )
    $images = @($Files)
    if ($images.Count -eq 0) { throw ($Context + '：没有图片。') }
    foreach ($image in $images) {
        if ($image.Length -eq 0) { throw ($Context + '：图片是空文件：' + $image.Name) }
    }

    if (@($images | Where-Object { $_.BaseName -notmatch '^\d+$' }).Count -eq 0) {
        return [pscustomobject]@{ Images = @(Get-NumericImages -DirectoryPath $images[0].DirectoryName -IgnoreCover); Mode = 'Numeric'; Warning = '' }
    }

    $records = @()
    foreach ($image in $images) {
        $record = Get-CompositeImageRecord -File $image
        if ($null -eq $record) {
            return [pscustomobject]@{
                Images = @($images | Sort-Object { Get-NaturalNameSortKey -Name $_.Name }, Name)
                Mode = 'NameSorted'
                Warning = ($Context + '：无法可靠提取连续页码，已按文件名称自然排序；这种命名方式无法自动判断是否缺图。')
            }
        }
        $records += $record
    }

    $orderedFiles = @()
    foreach ($group in @($records | Group-Object Prefix | Sort-Object {
        Get-NaturalNameSortKey -Name (Get-SequenceDisplayName -Prefix $_.Name)
    }, Name)) {
        $display = Get-SequenceDisplayName -Prefix $group.Name
        $duplicates = @($group.Group | Group-Object Page | Where-Object Count -gt 1)
        if ($duplicates.Count -gt 0) {
            $duplicate = $duplicates[0]
            throw ('{0}：分组 {1} 的页码 {2} 重复（{3}）' -f $Context, $display, $duplicate.Name, (($duplicate.Group.File.Name) -join '、'))
        }
        $numbers = @($group.Group.Page | Sort-Object -Unique)
        if ($numbers[0] -notin @([int64]0, [int64]1)) {
            throw ('{0}：分组 {1} 应从 000 或 001 开始，实际从 {2} 开始。' -f $Context, $display, $numbers[0])
        }
        $map = @{}
        foreach ($number in $numbers) { $map[[string]$number] = $true }
        for ($number = $numbers[0]; $number -le $numbers[-1]; $number++) {
            if (-not $map.ContainsKey([string]$number)) {
                throw ('{0}：分组 {1} 缺少页码 {2}。' -f $Context, $display, $number)
            }
        }
        $orderedFiles += @($group.Group | Sort-Object Page, @{ Expression = { $_.File.Name } } | ForEach-Object { $_.File })
    }
    return [pscustomobject]@{ Images = @($orderedFiles); Mode = 'Composite'; Warning = '' }
}

function Get-RootImageLayout {
    param([string]$ComicPath)
    $images = @(Get-RootBodyImageFiles -ComicPath $ComicPath)
    $standaloneZero = @($images | Where-Object { $_.BaseName -match '^0+$' })
    $records = @()
    $unrecognized = @()
    foreach ($image in $images) {
        if ($image.BaseName -match '^0+$') { continue }
        $record = Get-CompositeImageRecord -File $image
        if ($null -eq $record) { $unrecognized += $image } else { $records += $record }
    }
    $groups = @($records | Group-Object Prefix)
    if ($records.Count -eq 0 -or $unrecognized.Count -gt 0 -or $standaloneZero.Count -gt 1 -or $groups.Count -lt 2) {
        return [pscustomobject]@{ Recognized = $false; Entries = @(); CoverCandidate = $null }
    }
    $usedNames = @{}
    $entries = @()
    foreach ($group in @($groups | Sort-Object {
        Get-NaturalNameSortKey -Name (Get-SequenceDisplayName -Prefix $_.Name)
    }, Name)) {
        $display = Get-SequenceDisplayName -Prefix $group.Name
        $baseDisplay = $display
        $suffix = 2
        while ($usedNames.ContainsKey($display.ToLowerInvariant())) {
            $display = $baseDisplay + ' (' + $suffix + ')'
            $suffix++
        }
        $usedNames[$display.ToLowerInvariant()] = $true
        $sequence = Get-FlexibleImageSequence -Files @($group.Group.File) -Context $display
        $entries += [pscustomobject]@{
            Name = $display
            IsRootChapter = $true
            Images = @($sequence.Images)
            OrderMode = $sequence.Mode
            Warning = $sequence.Warning
        }
    }
    return [pscustomobject]@{
        Recognized = $true
        Entries = @($entries | Sort-Object { Get-NaturalNameSortKey -Name $_.Name }, Name)
        CoverCandidate = if ($standaloneZero.Count -eq 1) { $standaloneZero[0] } else { $null }
    }
}

function Get-PreferredCoverFile {
    param([string]$ComicPath)
    $namedCovers = @(Get-ChildItem -LiteralPath $ComicPath -File -ErrorAction SilentlyContinue | Where-Object {
        $_.BaseName -ieq 'cover' -and $script:ImageExtensions -contains $_.Extension.ToLowerInvariant() -and $_.Length -gt 0
    } | Sort-Object Name)
    if ($namedCovers.Count -gt 0) { return $namedCovers[0] }
    $layout = Get-RootImageLayout -ComicPath $ComicPath
    if ($layout.Recognized -and $null -ne $layout.CoverCandidate -and $layout.CoverCandidate.Length -gt 0) {
        return $layout.CoverCandidate
    }
    return $null
}

function Get-ChapterDirectories {
    param([string]$ComicPath)
    return @(Get-ChildItem -LiteralPath $ComicPath -Directory | Where-Object {
        $_.Name -ine $script:ReaderResourceFolderName -and
        @(Get-ChildItem -LiteralPath $_.FullName -File -ErrorAction SilentlyContinue | Where-Object {
            $script:ImageExtensions -contains $_.Extension.ToLowerInvariant() -and $_.BaseName -ine 'cover'
        }).Count -gt 0
    } | Sort-Object { Get-NaturalNameSortKey -Name $_.Name }, Name)
}

function Get-SourceChapterEntries {
    param([string]$ComicPath)
    $chapterDirectories = @(Get-ChapterDirectories -ComicPath $ComicPath)
    $rootImages = @(Get-RootBodyImageFiles -ComicPath $ComicPath)
    if ($chapterDirectories.Count -gt 0 -and $rootImages.Count -gt 0) {
        throw '漫画根目录中有正文图片，同时又存在章节文件夹；结构有歧义，请只保留一种正文结构。'
    }
    if ($chapterDirectories.Count -gt 0) {
        $entries = @($chapterDirectories | ForEach-Object {
            $files = @(Get-ChildItem -LiteralPath $_.FullName -File | Where-Object { $script:ImageExtensions -contains $_.Extension.ToLowerInvariant() })
            $sequence = Get-FlexibleImageSequence -Files $files -Context $_.Name
            [pscustomobject]@{
                Name = $_.Name
                IsRootChapter = $false
                Images = @($sequence.Images)
                OrderMode = $sequence.Mode
                Warning = $sequence.Warning
            }
        })
        return @(Set-SourceChapterEntryOrder -ComicPath $ComicPath -Entries $entries)
    }
    if ($rootImages.Count -gt 0) {
        $layout = Get-RootImageLayout -ComicPath $ComicPath
        if ($layout.Recognized) { return @(Set-SourceChapterEntryOrder -ComicPath $ComicPath -Entries @($layout.Entries)) }
        $sequence = Get-FlexibleImageSequence -Files $rootImages -Context $script:RootChapterToken
        $entries = @([pscustomobject]@{
            Name = $script:RootChapterToken
            IsRootChapter = $true
            Images = @($sequence.Images)
            OrderMode = $sequence.Mode
            Warning = $sequence.Warning
        })
        return @(Set-SourceChapterEntryOrder -ComicPath $ComicPath -Entries $entries)
    }
    return @()
}

function Get-SourceChapterInfo {
    param(
        [string]$LibraryRoot,
        [string]$SourceFolder,
        [string]$SourceChapter,
        [hashtable]$SourceEntriesCache = $null
    )
    if (-not (Test-SimpleFolderName $SourceFolder)) {
        throw ('来源漫画文件夹名称不合法：' + $SourceFolder)
    }
    $comicPath = Join-Path $LibraryRoot $SourceFolder
    if (-not (Test-Path -LiteralPath $comicPath -PathType Container)) {
        throw ('来源漫画文件夹不存在：' + $SourceFolder)
    }
    if ($null -ne $SourceEntriesCache -and $SourceEntriesCache.ContainsKey($SourceFolder)) {
        $entries = @($SourceEntriesCache[$SourceFolder])
    }
    else {
        $entries = @(Get-SourceChapterEntries -ComicPath $comicPath)
        if ($null -ne $SourceEntriesCache) { $SourceEntriesCache[$SourceFolder] = @($entries) }
    }
    $entry = @($entries | Where-Object { $_.Name -ceq $SourceChapter } | Select-Object -First 1)
    if ($entry.Count -eq 0) { throw ('来源章节不存在或命名分组已经变化：' + $SourceFolder + '\' + $SourceChapter) }
    return [pscustomobject]@{
        SourceFolder = $SourceFolder
        SourceChapter = $SourceChapter
        ComicPath = $comicPath
        ChapterPath = if ($entry[0].IsRootChapter) { $comicPath } else { Join-Path $comicPath $SourceChapter }
        IsRootChapter = $entry[0].IsRootChapter
        Images = @($entry[0].Images)
        Count = @($entry[0].Images).Count
        OrderMode = $entry[0].OrderMode
        Warning = $entry[0].Warning
        HasConfiguredCover = ($null -ne $entry[0].PSObject.Properties['ConfiguredCoverMode'])
        ConfiguredCoverMode = if ($null -ne $entry[0].PSObject.Properties['ConfiguredCoverMode']) { [string]$entry[0].ConfiguredCoverMode } else { 'first' }
        ConfiguredCoverFile = if ($null -ne $entry[0].PSObject.Properties['ConfiguredCoverFile']) { [string]$entry[0].ConfiguredCoverFile } else { '' }
        ConfiguredShowChapterCovers = if ($null -ne $entry[0].PSObject.Properties['ConfiguredShowChapterCovers']) { [bool]$entry[0].ConfiguredShowChapterCovers } else { $false }
    }
}

function Resolve-ConfiguredSourceCoverPath {
    param(
        [string]$ComicPath,
        [string]$RelativePath
    )
    if ([string]::IsNullOrWhiteSpace($RelativePath) -or [IO.Path]::IsPathRooted($RelativePath)) { return '' }
    try {
        $rootFull = [IO.Path]::GetFullPath($ComicPath).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
        $candidate = [IO.Path]::GetFullPath((Join-Path $rootFull ($RelativePath -replace '/', [IO.Path]::DirectorySeparatorChar)))
        $prefix = $rootFull + [IO.Path]::DirectorySeparatorChar
        if (-not $candidate.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return '' }
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { return '' }
        $file = Get-Item -LiteralPath $candidate
        if ($script:ImageExtensions -notcontains $file.Extension.ToLowerInvariant() -or $file.Length -eq 0) { return '' }
        return $file.FullName
    }
    catch { return '' }
}

function Get-CandidateFolders {
    param([string]$LibraryRoot)
    $candidates = @()
    foreach ($directory in @(Get-ChildItem -LiteralPath $LibraryRoot -Directory | Sort-Object Name)) {
        if ($directory.Name -eq $script:OutputFolderName -or $directory.Name -ieq $script:ReaderResourceFolderName) { continue }
        $chapterDirectories = @(Get-ChapterDirectories -ComicPath $directory.FullName)
        $rootImages = @(Get-RootBodyImageFiles -ComicPath $directory.FullName)
        if ($chapterDirectories.Count -eq 0 -and $rootImages.Count -eq 0) { continue }
        $candidates += [pscustomobject]@{
            Name = $directory.Name
            FullName = $directory.FullName
            ChapterCount = if ($chapterDirectories.Count -gt 0) { $chapterDirectories.Count } else { @(Get-SourceChapterEntries -ComicPath $directory.FullName).Count }
            IsAmbiguous = ($chapterDirectories.Count -gt 0 -and $rootImages.Count -gt 0)
            UsesRootImages = ($chapterDirectories.Count -eq 0 -and $rootImages.Count -gt 0)
            HasCover = ($null -ne (Get-PreferredCoverFile -ComicPath $directory.FullName))
        }
    }
    return @($candidates)
}

function Read-OrganizerPlan {
    param([string]$Path)
    $raw = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    return ($raw | ConvertFrom-Json)
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
    $lines = @($text -replace "`r`n?", "`n" -split "`n" | ForEach-Object {
        ([regex]::Replace($_, '[\t ]+', ' ')).Trim()
    })
    $normalized = New-Object 'System.Collections.Generic.List[string]'
    $lastWasBlank = $true
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            if (-not $lastWasBlank) { $normalized.Add('') }
            $lastWasBlank = $true
        }
        else {
            $normalized.Add($line)
            $lastWasBlank = $false
        }
    }
    while ($normalized.Count -gt 0 -and [string]::IsNullOrWhiteSpace($normalized[$normalized.Count - 1])) { $normalized.RemoveAt($normalized.Count - 1) }
    return ($normalized -join "`r`n").Trim()
}

function Get-MetadataDescription {
    param([AllowNull()][object]$Metadata)
    foreach ($fieldName in @('description', 'intro', 'summary', 'desc', '简介', '簡介')) {
        $value = Get-ObjectProperty -Object $Metadata -Name $fieldName -Default $null
        $text = ConvertTo-MetadataPlainText -Value $value
        if (-not [string]::IsNullOrWhiteSpace($text)) { return $text }
    }
    return ''
}

function Get-SourceDescription {
    param(
        [string]$LibraryRoot,
        [string]$SourceFolder
    )
    if ([string]::IsNullOrWhiteSpace($SourceFolder)) { return '' }
    $metadataPath = Join-Path (Join-Path $LibraryRoot $SourceFolder) '元数据.json'
    if (-not (Test-Path -LiteralPath $metadataPath -PathType Leaf)) { return '' }
    try {
        $metadata = [IO.File]::ReadAllText($metadataPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        return Get-MetadataDescription -Metadata $metadata
    }
    catch {
        return ''
    }
}

function Test-OrganizerPlan {
    param(
        [object]$Plan,
        [string]$LibraryRoot,
        [switch]$IgnoreExistingOutput,
        [string]$ProgressPrefix = '正在核验方案'
    )

    $errors = New-Object 'System.Collections.Generic.List[string]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    $omissionWarnings = New-Object 'System.Collections.Generic.List[string]'
    $overlapWarnings = New-Object 'System.Collections.Generic.List[string]'
    $simpleOverlapWarnings = New-Object 'System.Collections.Generic.List[string]'
    $outputName = [string](Get-ObjectProperty -Object $Plan -Name 'outputName' -Default '')
    if (-not (Test-SimpleFolderName $outputName)) {
        $errors.Add('输出漫画名称为空或含有 Windows 文件夹不允许的字符。')
    }
    elseif ($outputName -eq $script:OutputFolderName) {
        $errors.Add(('输出漫画名称不能是“' + $script:OutputFolderName + '”。'))
    }

    $outputBase = [IO.Path]::GetFullPath((Join-Path $LibraryRoot $script:OutputFolderName))
    $outputPath = ''
    if (Test-SimpleFolderName $outputName) {
        $outputPath = [IO.Path]::GetFullPath((Join-Path $outputBase $outputName))
        if ([IO.Path]::GetDirectoryName($outputPath) -cne $outputBase) {
            $errors.Add('输出路径超出了整理完成目录。')
        }
        elseif (-not $IgnoreExistingOutput -and (Test-Path -LiteralPath $outputPath)) {
            $errors.Add(('输出文件夹已经存在，不会覆盖：' + $outputPath))
        }
    }

    $planChapters = @((Get-ObjectProperty -Object $Plan -Name 'chapters' -Default @()))
    $showChapterCovers = [bool](Get-ObjectProperty -Object $Plan -Name 'showChapterCovers' -Default $false)
    if ($planChapters.Count -eq 0) {
        $errors.Add('方案中没有章节。')
    }

    $resolvedChapters = @()
    $sourceCache = @{}
    $sourceEntriesCache = @{}
    $usageBySource = @{}
    $rangeHistoryBySource = @{}
    $selectedSourceFolders = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($sourceFolderValue in @((Get-ObjectProperty -Object $Plan -Name 'selectedSourceFolders' -Default @()))) {
        $sourceFolderText = ([string]$sourceFolderValue).Trim()
        if (-not (Test-SimpleFolderName -Name $sourceFolderText)) {
            $errors.Add(('所选来源漫画文件夹名称无效：' + $sourceFolderText))
            continue
        }
        $sourceFolderPath = Join-Path $LibraryRoot $sourceFolderText
        if (-not (Test-Path -LiteralPath $sourceFolderPath -PathType Container)) {
            $errors.Add(('所选来源漫画文件夹不存在：' + $sourceFolderText))
            continue
        }
        [void]$selectedSourceFolders.Add($sourceFolderText)
    }
    $planRowIndex = 0

    foreach ($chapter in $planChapters) {
        $planRowIndex++
        Update-OrganizerProgress -Message ('{0}：第 {1}/{2} 行' -f $ProgressPrefix, $planRowIndex, $planChapters.Count)
        $mergeWithPrevious = [bool](Get-ObjectProperty -Object $chapter -Name 'mergeWithPrevious' -Default $false)
        $chapterCoverMode = ([string](Get-ObjectProperty -Object $chapter -Name 'chapterCoverMode' -Default 'first')).Trim().ToLowerInvariant()
        $chapterCoverSourcePath = ([string](Get-ObjectProperty -Object $chapter -Name 'chapterCoverPath' -Default '')).Trim()
        $chapterCoverFile = ''
        if ($chapterCoverMode -notin @('metadata', 'first', 'custom', 'none')) { $chapterCoverMode = 'first' }
        if ($mergeWithPrevious) {
            if ($chapterCoverMode -eq 'custom' -and -not [string]::IsNullOrWhiteSpace($chapterCoverSourcePath)) {
                $warnings.Add(('第 {0} 行已并入上一话，其章节封面设置不会单独使用。' -f $planRowIndex))
            }
        }
        $labelInfo = $null
        $number = ''
        $chapterLabel = ''
        $title = ''
        $folderName = ''
        if ($mergeWithPrevious) {
            if ($resolvedChapters.Count -eq 0) {
                $errors.Add(('第 ' + $planRowIndex + ' 行不能并入上一话，因为前面还没有输出章节。'))
                continue
            }
            $targetChapter = $resolvedChapters[-1]
            $number = $targetChapter.Number
            $chapterLabel = $targetChapter.DisplayLabel
            $title = $targetChapter.Title
            $folderName = $targetChapter.FolderName
        }
        else {
            $numberValue = Get-ObjectProperty -Object $chapter -Name 'number' -Default 0
            $labelInfo = ConvertTo-ChapterLabelInfo -Value $numberValue
            if ($null -eq $labelInfo) {
                $errors.Add(('章节话序/特殊标签无效；可填 4.5、特典话、番外篇或插画集：' + [string]$numberValue))
                continue
            }
            $number = $labelInfo.Number
            $chapterLabel = $labelInfo.DisplayLabel
            $title = [string](Get-ObjectProperty -Object $chapter -Name 'title' -Default '')
            if ($title -match '[<>:"/\\|?*]' -or $title.EndsWith(' ') -or $title.EndsWith('.')) {
                $errors.Add(($chapterLabel + '的章节名含有文件夹不允许的字符。'))
            }
            $folderName = $labelInfo.BaseFolderName
            if (-not [string]::IsNullOrWhiteSpace($title)) { $folderName += ' ' + $title.Trim() }
            if (-not (Test-SimpleFolderName -Name $folderName)) {
                $errors.Add(($chapterLabel + '生成的章节文件夹名称无效。'))
            }
        }
        $sourceFolder = [string](Get-ObjectProperty -Object $chapter -Name 'sourceFolder' -Default '')
        $sourceChapter = [string](Get-ObjectProperty -Object $chapter -Name 'sourceChapter' -Default '')
        $start = 0
        $end = 0
        if (-not [int]::TryParse([string](Get-ObjectProperty -Object $chapter -Name 'start' -Default 0), [ref]$start)) {
            $errors.Add(($chapterLabel + '的起始序号不是整数。'))
            continue
        }
        if (-not [int]::TryParse([string](Get-ObjectProperty -Object $chapter -Name 'end' -Default 0), [ref]$end)) {
            $errors.Add(($chapterLabel + '的结束序号不是整数。'))
            continue
        }
        if ($start -lt 1 -or $end -lt $start) {
            $errors.Add(($chapterLabel + '的范围无效：' + $start + '-' + $end))
            continue
        }

        $sourceKey = $sourceFolder + '|' + $sourceChapter
        if (-not $sourceCache.ContainsKey($sourceKey)) {
            try {
                $sourceCache[$sourceKey] = Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $sourceFolder -SourceChapter $sourceChapter -SourceEntriesCache $sourceEntriesCache
            }
            catch {
                $errors.Add(($sourceFolder + '\' + $sourceChapter + '：' + $_.Exception.Message))
                continue
            }
        }
        $sourceInfo = $sourceCache[$sourceKey]
        if (-not [string]::IsNullOrWhiteSpace($sourceInfo.Warning) -and -not $warnings.Contains($sourceInfo.Warning)) {
            $warnings.Add($sourceInfo.Warning)
        }
        if ($end -gt $sourceInfo.Count) {
            $errors.Add(('{0}的范围 {1}-{2} 超出来源总数 {3}：{4}\{5}' -f $chapterLabel, $start, $end, $sourceInfo.Count, $sourceFolder, $sourceChapter))
            continue
        }

        if (-not $mergeWithPrevious -and $chapterCoverMode -eq 'metadata') {
            $chapterCoverMode = [string]$sourceInfo.ConfiguredCoverMode
            if ($chapterCoverMode -notin @('first', 'custom', 'none')) { $chapterCoverMode = 'first' }
            if ($chapterCoverMode -eq 'custom') {
                $chapterCoverSourcePath = Resolve-ConfiguredSourceCoverPath -ComicPath $sourceInfo.ComicPath -RelativePath ([string]$sourceInfo.ConfiguredCoverFile)
                if ([string]::IsNullOrWhiteSpace($chapterCoverSourcePath)) {
                    $warnings.Add(('第 {0} 行无法采用来源元数据中的自选章节封面，已回退为首图：{1}\{2}' -f $planRowIndex, $sourceFolder, $sourceChapter))
                    $chapterCoverMode = 'first'
                }
            }
        }
        if (-not $mergeWithPrevious -and $chapterCoverMode -eq 'custom') {
            if (-not (Test-Path -LiteralPath $chapterCoverSourcePath -PathType Leaf)) {
                $errors.Add(('第 {0} 行选择的章节封面不存在：{1}' -f $planRowIndex, $chapterCoverSourcePath))
            }
            else {
                $chapterCoverSourceFile = Get-Item -LiteralPath $chapterCoverSourcePath
                if ($script:ImageExtensions -notcontains $chapterCoverSourceFile.Extension.ToLowerInvariant() -or $chapterCoverSourceFile.Length -eq 0) {
                    $errors.Add(('第 {0} 行选择的章节封面不是有效图片：{1}' -f $planRowIndex, $chapterCoverSourcePath))
                }
                else {
                    $chapterCoverFile = $script:ReaderResourceFolderName + '/' + $script:ChapterCoverFolderName + '/chapter-' + (($resolvedChapters.Count + 1).ToString('D4')) + $chapterCoverSourceFile.Extension.ToLowerInvariant()
                }
            }
        }

        if (-not $usageBySource.ContainsKey($sourceKey)) {
            $usageBySource[$sourceKey] = @{}
        }
        if (-not $rangeHistoryBySource.ContainsKey($sourceKey)) { $rangeHistoryBySource[$sourceKey] = @() }
        $usage = $usageBySource[$sourceKey]
        $overlapNumbers = New-Object 'System.Collections.Generic.List[int]'
        $overlapLabels = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        for ($imageNumber = $start; $imageNumber -le $end; $imageNumber++) {
            if ($imageNumber -eq $start -or $imageNumber -eq $end -or (($imageNumber - $start) % 250) -eq 0) {
                Update-OrganizerProgress -Message ('{0}：第 {1}/{2} 行｜检查图片 {3}/{4}' -f $ProgressPrefix, $planRowIndex, $planChapters.Count, ($imageNumber - $start + 1), ($end - $start + 1))
            }
            if ($usage.ContainsKey([string]$imageNumber)) {
                $overlapNumbers.Add($imageNumber)
                [void]$overlapLabels.Add([string]$usage[[string]$imageNumber])
            }
            else {
                $usage[[string]$imageNumber] = $chapterLabel
            }
        }
        $hasOverlap = $overlapNumbers.Count -gt 0
        if ($hasOverlap) {
            $overlapRangeText = ConvertTo-NumberRangeText -Numbers @($overlapNumbers)
            $previousText = @($overlapLabels) -join '、'
            $overlapMessage = ('图片被重复使用：{0}\{1}（{2}；本行“{3}”与前面“{4}”重叠）' -f $sourceFolder, $sourceChapter, $overlapRangeText, $chapterLabel, $previousText)
            $overlapWarnings.Add($overlapMessage)
            $warnings.Add($overlapMessage)
            $history = @($rangeHistoryBySource[$sourceKey])
            $simpleBoundary = $overlapNumbers.Count -eq 1 -and $overlapNumbers[0] -eq $start -and @($history | Where-Object { [int]$_.End -eq $start }).Count -gt 0
            if ($simpleBoundary) {
                $simpleMessage = ('疑似相邻范围边界写重：{0}\{1} 的 {2} 被上一段和本段同时使用；本段起始很可能应为 {3}。' -f $sourceFolder, $sourceChapter, $start, ($start + 1))
                $simpleOverlapWarnings.Add($simpleMessage)
            }
        }
        $rangeHistoryBySource[$sourceKey] = @($rangeHistoryBySource[$sourceKey]) + @([pscustomobject]@{ Start = $start; End = $end; Label = $chapterLabel; Row = $planRowIndex })
        [void]$selectedSourceFolders.Add($sourceFolder)
        $selectedImages = @($sourceInfo.Images[($start - 1)..($end - 1)])
        if ($mergeWithPrevious) {
            $targetChapter = $resolvedChapters[-1]
            $targetChapter.Images = @($targetChapter.Images) + @($selectedImages)
            $targetChapter.SourceParts = @($targetChapter.SourceParts) + @([pscustomobject]@{
                SourceFolder = $sourceFolder
                SourceChapter = $sourceChapter
                Start = $start
                End = $end
            })
            $targetChapter.HasOverlap = ($targetChapter.HasOverlap -or $hasOverlap)
        }
        else {
            $resolvedChapters += [pscustomobject]@{
                IsNumeric = $labelInfo.IsNumeric
                Number = $number
                SortNumber = $labelInfo.SortNumber
                ReadingOrder = $resolvedChapters.Count + 1
                DisplayLabel = $chapterLabel
                FolderName = $folderName
                Title = $title.Trim()
                SourceFolder = $sourceFolder
                SourceChapter = $sourceChapter
                Start = $start
                End = $end
                SourceTotal = $sourceInfo.Count
                SourceParts = @([pscustomobject]@{ SourceFolder = $sourceFolder; SourceChapter = $sourceChapter; Start = $start; End = $end })
                Images = @($selectedImages)
                HasOverlap = $hasOverlap
                ChapterCoverMode = $chapterCoverMode
                ChapterCoverSourcePath = $chapterCoverSourcePath
                ChapterCoverFile = $chapterCoverFile
            }
        }
    }

    $duplicates = @($resolvedChapters | Where-Object IsNumeric | Group-Object Number | Where-Object Count -gt 1)
    foreach ($duplicate in $duplicates) {
        $errors.Add(('输出章节编号重复：第 ' + $duplicate.Name + ' 话。'))
    }
    $folderDuplicates = @($resolvedChapters | Group-Object FolderName | Where-Object Count -gt 1)
    foreach ($duplicate in $folderDuplicates) {
        $errors.Add(('输出章节文件夹名称重复：' + $duplicate.Name))
    }
    $numericChapters = @($resolvedChapters | Where-Object IsNumeric | Sort-Object SortNumber)
    if ($numericChapters.Count -gt 0) {
        if ($numericChapters[0].SortNumber -ne [decimal]1) {
            $errors.Add(('数字主线章节必须从第 1 话开始，实际从第 ' + $numericChapters[0].Number + ' 话开始。'))
        }
        $numberMap = @{}
        foreach ($chapter in $numericChapters) { $numberMap[$chapter.Number] = $true }
        $maxWholeChapter = [int][decimal]::Floor($numericChapters[-1].SortNumber)
        for ($wholeNumber = 1; $wholeNumber -le $maxWholeChapter; $wholeNumber++) {
            $integerKey = $wholeNumber.ToString([Globalization.CultureInfo]::InvariantCulture)
            if (-not $numberMap.ContainsKey($integerKey)) {
                $errors.Add(('输出章节缺少第 ' + $wholeNumber + ' 话。'))
            }
        }
    }

    $selectedSourceArray = @($selectedSourceFolders)
    for ($sourceFolderIndex = 0; $sourceFolderIndex -lt $selectedSourceArray.Count; $sourceFolderIndex++) {
        $sourceFolder = $selectedSourceArray[$sourceFolderIndex]
        Update-OrganizerProgress -Message ('{0}：检查来源 {1}/{2}｜{3}' -f $ProgressPrefix, ($sourceFolderIndex + 1), $selectedSourceArray.Count, $sourceFolder)
        $comicPath = Join-Path $LibraryRoot $sourceFolder
        try {
            if ($sourceEntriesCache.ContainsKey($sourceFolder)) {
                $sourceChapterEntries = @($sourceEntriesCache[$sourceFolder])
            }
            else {
                $sourceChapterEntries = @(Get-SourceChapterEntries -ComicPath $comicPath)
                $sourceEntriesCache[$sourceFolder] = @($sourceChapterEntries)
            }
        }
        catch {
            $errors.Add(($sourceFolder + '：' + $_.Exception.Message))
            continue
        }
        for ($sourceChapterIndex = 0; $sourceChapterIndex -lt $sourceChapterEntries.Count; $sourceChapterIndex++) {
            $chapterDirectory = $sourceChapterEntries[$sourceChapterIndex]
            Update-OrganizerProgress -Message ('{0}：检查来源 {1}/{2}｜章节 {3}/{4}：{5}' -f $ProgressPrefix, ($sourceFolderIndex + 1), $selectedSourceArray.Count, ($sourceChapterIndex + 1), $sourceChapterEntries.Count, $chapterDirectory.Name)
            $sourceKey = $sourceFolder + '|' + $chapterDirectory.Name
            try {
                if (-not $sourceCache.ContainsKey($sourceKey)) {
                    $sourceCache[$sourceKey] = Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $sourceFolder -SourceChapter $chapterDirectory.Name -SourceEntriesCache $sourceEntriesCache
                }
                $sourceInfo = $sourceCache[$sourceKey]
                if (-not $usageBySource.ContainsKey($sourceKey)) {
                    $message = ('来源章节完全未选择：{0}\{1}（{2} 张图片不会输出）' -f $sourceFolder, $chapterDirectory.Name, $sourceInfo.Count)
                    $omissionWarnings.Add($message)
                    $warnings.Add($message)
                    continue
                }
                $usage = $usageBySource[$sourceKey]
                $unusedNumbers = New-Object 'System.Collections.Generic.List[int]'
                for ($imageNumber = 1; $imageNumber -le $sourceInfo.Count; $imageNumber++) {
                    if (-not $usage.ContainsKey([string]$imageNumber)) {
                        $unusedNumbers.Add($imageNumber)
                    }
                }
                if ($unusedNumbers.Count -gt 0) {
                    $rangeText = ConvertTo-NumberRangeText -Numbers @($unusedNumbers)
                    $message = ('有 {0} 张来源图片未选择：{1}\{2}（{3}）' -f $unusedNumbers.Count, $sourceFolder, $chapterDirectory.Name, $rangeText)
                    $omissionWarnings.Add($message)
                    $warnings.Add($message)
                }
            }
            catch {
                $errors.Add(($sourceFolder + '\' + $chapterDirectory.Name + '：' + $_.Exception.Message))
            }
        }
    }

    $coverSource = [string](Get-ObjectProperty -Object $Plan -Name 'coverSource' -Default '')
    if ([string]::IsNullOrWhiteSpace($coverSource) -and $resolvedChapters.Count -gt 0) {
        $coverSource = $resolvedChapters[0].SourceFolder
    }
    $coverPath = ''
    $coverOutputName = 'cover.jpg'
    $coverIsAutomatic = $false
    $customCoverPath = ([string](Get-ObjectProperty -Object $Plan -Name 'customCoverPath' -Default '')).Trim()
    $coverIsCustom = $false
    if (-not [string]::IsNullOrWhiteSpace($customCoverPath)) {
        if (-not (Test-Path -LiteralPath $customCoverPath -PathType Leaf)) {
            $errors.Add(('自选漫画封面不存在：' + $customCoverPath))
        }
        else {
            $customCoverFile = Get-Item -LiteralPath $customCoverPath
            if ($script:ImageExtensions -notcontains $customCoverFile.Extension.ToLowerInvariant() -or $customCoverFile.Length -eq 0) {
                $errors.Add(('自选漫画封面不是有效图片：' + $customCoverPath))
            }
            else {
                $coverPath = $customCoverFile.FullName
                $coverOutputName = 'cover' + $customCoverFile.Extension.ToLowerInvariant()
                $coverIsCustom = $true
            }
        }
    }
    elseif (-not (Test-SimpleFolderName $coverSource)) {
        $errors.Add('封面来源文件夹无效。')
    }
    else {
        $coverComicPath = Join-Path $LibraryRoot $coverSource
        $preferredCover = Get-PreferredCoverFile -ComicPath $coverComicPath
        if ($null -ne $preferredCover) {
            $coverPath = $preferredCover.FullName
            $coverOutputName = 'cover' + $preferredCover.Extension.ToLowerInvariant()
            if ($preferredCover.BaseName -ine 'cover') {
                $coverIsAutomatic = $true
                $warnings.Add(('已将根目录中独立的零号图片作为封面：{0}\{1}' -f $coverSource, $preferredCover.Name))
            }
        }
        else {
            try {
                if ($sourceEntriesCache.ContainsKey($coverSource)) {
                    $coverChapterDirectories = @($sourceEntriesCache[$coverSource])
                }
                else {
                    $coverChapterDirectories = @(Get-SourceChapterEntries -ComicPath $coverComicPath)
                    $sourceEntriesCache[$coverSource] = @($coverChapterDirectories)
                }
                if ($coverChapterDirectories.Count -eq 0) {
                    throw '没有可识别的章节文件夹。'
                }
                $coverChapterInfo = Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $coverSource -SourceChapter $coverChapterDirectories[0].Name -SourceEntriesCache $sourceEntriesCache
                $coverImage = $coverChapterInfo.Images[0]
                $coverPath = $coverImage.FullName
                $coverOutputName = 'cover' + $coverImage.Extension.ToLowerInvariant()
                $coverIsAutomatic = $true
                $warnings.Add(('封面来源没有有效的 cover.jpg，已自动使用第一章首图并输出为 {0}：{1}\{2}\{3}' -f $coverOutputName, $coverSource, $coverChapterDirectories[0].Name, $coverImage.Name))
            }
            catch {
                $errors.Add(('封面来源缺少有效 cover.jpg，且无法取得第一章 0001：{0}（{1}）' -f $coverSource, $_.Exception.Message))
            }
        }
        if ($selectedSourceFolders.Count -gt 0 -and -not $selectedSourceFolders.Contains($coverSource)) {
            $warnings.Add('封面来源不在章节来源中，但仍可使用。')
        }
    }

    $descriptionProperty = $Plan.PSObject.Properties['description']
    $descriptionSource = [string](Get-ObjectProperty -Object $Plan -Name 'descriptionSource' -Default '')
    $description = if ($null -ne $descriptionProperty) { [string]$descriptionProperty.Value } else { '' }
    if ($null -eq $descriptionProperty) {
        # 兼容旧版方案：旧方案没有简介字段时，沿用原先从封面来源继承简介的行为。
        if ([string]::IsNullOrWhiteSpace($descriptionSource)) { $descriptionSource = $coverSource }
        $description = Get-SourceDescription -LibraryRoot $LibraryRoot -SourceFolder $descriptionSource
    }
    if (-not [string]::IsNullOrWhiteSpace($descriptionSource)) {
        if (-not (Test-SimpleFolderName -Name $descriptionSource)) {
            $errors.Add('简介来源文件夹名称无效。')
        }
        else {
            $descriptionSourcePath = Join-Path $LibraryRoot $descriptionSource
            if (-not (Test-Path -LiteralPath $descriptionSourcePath -PathType Container)) {
                $errors.Add(('简介来源文件夹不存在：' + $descriptionSource))
            }
            elseif ($selectedSourceFolders.Count -gt 0 -and -not $selectedSourceFolders.Contains($descriptionSource)) {
                $warnings.Add('简介来源不在章节来源中，但已保留手动编辑后的简介文本。')
            }
        }
    }

    $totalImages = 0
    foreach ($chapter in $resolvedChapters) { $totalImages += $chapter.Images.Count }
    Update-OrganizerProgress -Message ('{0}完成：{1} 话，{2} 张图片' -f $ProgressPrefix, $resolvedChapters.Count, $totalImages)
    return [pscustomobject]@{
        IsValid = ($errors.Count -eq 0)
        Errors = @($errors)
        Warnings = @($warnings)
        OmissionWarnings = @($omissionWarnings)
        OverlapWarnings = @($overlapWarnings)
        SimpleOverlapWarnings = @($simpleOverlapWarnings)
        OutputName = $outputName
        OutputBase = $outputBase
        OutputPath = $outputPath
        CoverSource = $coverSource
        CoverPath = $coverPath
        CoverOutputName = $coverOutputName
        CoverIsAutomatic = $coverIsAutomatic
        CoverIsCustom = $coverIsCustom
        Description = $description
        DescriptionSource = $descriptionSource
        ShowChapterCovers = $showChapterCovers
        Chapters = @($resolvedChapters | Sort-Object ReadingOrder)
        ChapterCount = $resolvedChapters.Count
        TotalImages = $totalImages
        SourceFolders = @($selectedSourceFolders)
        Plan = $Plan
    }
}

function New-OutputMetadata {
    param(
        [object]$Audit,
        [string]$LibraryRoot
    )
    $metadataSourcePath = Join-Path (Join-Path $LibraryRoot $Audit.CoverSource) '元数据.json'
    $metadata = $null
    if (Test-Path -LiteralPath $metadataSourcePath -PathType Leaf) {
        try {
            $metadata = [IO.File]::ReadAllText($metadataSourcePath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        }
        catch {
            $metadata = $null
        }
    }
    if ($null -eq $metadata) {
        $metadata = [pscustomobject][ordered]@{
            name = $Audit.OutputName
            author = @()
            description = ''
        }
    }

    $readingOrder = 0
    $chapterInfos = @($Audit.Chapters | ForEach-Object {
        $readingOrder++
        [pscustomobject][ordered]@{
            chapterTitle = $_.FolderName
            chapterFolder = $_.FolderName
            displayNumber = $_.Number
            displayLabel = $_.DisplayLabel
            order = $readingOrder
            coverMode = [string]$_.ChapterCoverMode
            coverFile = [string]$_.ChapterCoverFile
        }
    })
    $organizerInfo = [pscustomobject][ordered]@{
        schemaVersion = 6
        generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        sourceFolders = @($Audit.SourceFolders)
        chapterCount = $Audit.ChapterCount
        totalImages = $Audit.TotalImages
        coverFile = $Audit.CoverOutputName
        coverAutomatic = $Audit.CoverIsAutomatic
        coverCustom = $Audit.CoverIsCustom
        descriptionSource = $Audit.DescriptionSource
    }
    $metadata | Add-Member -NotePropertyName name -NotePropertyValue $Audit.OutputName -Force
    $metadata | Add-Member -NotePropertyName description -NotePropertyValue $Audit.Description -Force
    $metadata | Add-Member -NotePropertyName chapterInfos -NotePropertyValue $chapterInfos -Force
    $metadata | Add-Member -NotePropertyName readerOptions -NotePropertyValue ([pscustomobject][ordered]@{
        showChapterCovers = [bool]$Audit.ShowChapterCovers
    }) -Force
    $metadata | Add-Member -NotePropertyName organizer -NotePropertyValue $organizerInfo -Force
    return $metadata
}

function Test-NormalizedOutput {
    param(
        [string]$OutputPath,
        [object]$Audit,
        [string]$ProgressPrefix = '正在复核输出'
    )
    $errors = New-Object 'System.Collections.Generic.List[string]'
    $outputCoverPath = Join-Path $OutputPath $Audit.CoverOutputName
    if (-not (Test-Path -LiteralPath $outputCoverPath -PathType Leaf)) {
        $errors.Add(('输出缺少封面：' + $Audit.CoverOutputName))
    }
    elseif ((Get-Item -LiteralPath $outputCoverPath).Length -eq 0) {
        $errors.Add(('输出封面是空文件：' + $Audit.CoverOutputName))
    }
    for ($chapterIndex = 0; $chapterIndex -lt $Audit.Chapters.Count; $chapterIndex++) {
        $chapter = $Audit.Chapters[$chapterIndex]
        Update-OrganizerProgress -Message ('{0}：第 {1}/{2} 话｜{3}' -f $ProgressPrefix, ($chapterIndex + 1), $Audit.Chapters.Count, $chapter.FolderName)
        $folderName = $chapter.FolderName
        $chapterPath = Join-Path $OutputPath $folderName
        if (-not (Test-Path -LiteralPath $chapterPath -PathType Container)) {
            $errors.Add(('输出缺少章节：' + $folderName))
            continue
        }
        try {
            $images = @(Get-NumericImages -DirectoryPath $chapterPath)
            if ($images.Count -ne $chapter.Images.Count) {
                $errors.Add(('{0} 图片数量错误：应有 {1}，实际 {2}' -f $folderName, $chapter.Images.Count, $images.Count))
            }
        }
        catch {
            $errors.Add(($folderName + '：' + $_.Exception.Message))
        }
        if ([string]$chapter.ChapterCoverMode -eq 'custom') {
            $coverRelativePath = ([string]$chapter.ChapterCoverFile) -replace '/', [IO.Path]::DirectorySeparatorChar
            $customChapterCoverPath = Join-Path $OutputPath $coverRelativePath
            if (-not (Test-Path -LiteralPath $customChapterCoverPath -PathType Leaf)) {
                $errors.Add(($folderName + '：缺少自选章节封面 ' + $chapter.ChapterCoverFile))
            }
            elseif ((Get-Item -LiteralPath $customChapterCoverPath).Length -eq 0) {
                $errors.Add(($folderName + '：自选章节封面是空文件。'))
            }
        }
    }
    return @($errors)
}

function Invoke-OrganizerPlan {
    param(
        [object]$Audit,
        [string]$LibraryRoot,
        [string]$ProgressPrefix = '正在整理'
    )
    if (-not $Audit.IsValid) {
        throw '整理方案未通过核验。'
    }
    if (-not (Test-Path -LiteralPath $Audit.OutputBase -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $Audit.OutputBase)
    }
    if (Test-Path -LiteralPath $Audit.OutputPath) {
        throw ('输出文件夹已经存在，不会覆盖：' + $Audit.OutputPath)
    }

    $temporaryName = '.comic-organizer-' + [guid]::NewGuid().ToString('N')
    $temporaryPath = [IO.Path]::GetFullPath((Join-Path $Audit.OutputBase $temporaryName))
    $expectedPrefix = $Audit.OutputBase.TrimEnd('\') + '\.comic-organizer-'
    if (-not $temporaryPath.StartsWith($expectedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw '临时输出路径校验失败。'
    }

    try {
        [void](New-Item -ItemType Directory -Path $temporaryPath)
        Update-OrganizerProgress -Message ($ProgressPrefix + '：正在复制封面')
        Copy-Item -LiteralPath $Audit.CoverPath -Destination (Join-Path $temporaryPath $Audit.CoverOutputName)

        $copiedImages = 0
        for ($chapterIndex = 0; $chapterIndex -lt $Audit.Chapters.Count; $chapterIndex++) {
            $chapter = $Audit.Chapters[$chapterIndex]
            $folderName = $chapter.FolderName
            $chapterPath = Join-Path $temporaryPath $folderName
            [void](New-Item -ItemType Directory -Path $chapterPath)
            $outputNumber = 0
            foreach ($image in $chapter.Images) {
                $outputNumber++
                $copiedImages++
                Update-OrganizerProgress -Message ('{0}：第 {1}/{2} 话｜{3}｜图片 {4}/{5}｜总体 {6}/{7}' -f $ProgressPrefix, ($chapterIndex + 1), $Audit.Chapters.Count, $folderName, $outputNumber, $chapter.Images.Count, $copiedImages, $Audit.TotalImages)
                $destinationName = $outputNumber.ToString('D4') + $image.Extension.ToLowerInvariant()
                Copy-Item -LiteralPath $image.FullName -Destination (Join-Path $chapterPath $destinationName)
            }
        }

        $customChapterCovers = @($Audit.Chapters | Where-Object { [string]$_.ChapterCoverMode -eq 'custom' })
        if ($customChapterCovers.Count -gt 0) {
            Update-OrganizerProgress -Message ($ProgressPrefix + '：正在复制自选章节封面')
            foreach ($chapter in $customChapterCovers) {
                $relativePath = ([string]$chapter.ChapterCoverFile) -replace '/', [IO.Path]::DirectorySeparatorChar
                $destination = Join-Path $temporaryPath $relativePath
                $destinationDirectory = [IO.Path]::GetDirectoryName($destination)
                if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
                    [void](New-Item -ItemType Directory -Path $destinationDirectory -Force)
                }
                Copy-Item -LiteralPath ([string]$chapter.ChapterCoverSourcePath) -Destination $destination -Force
            }
        }

        Update-OrganizerProgress -Message ($ProgressPrefix + '：正在写入元数据与整理方案')
        $metadata = New-OutputMetadata -Audit $Audit -LibraryRoot $LibraryRoot
        Write-JsonFile -Path (Join-Path $temporaryPath '元数据.json') -Value $metadata
        Write-JsonFile -Path (Join-Path $temporaryPath '整理方案.json') -Value $Audit.Plan

        $verifyErrors = @(Test-NormalizedOutput -OutputPath $temporaryPath -Audit $Audit -ProgressPrefix '正在复核整理结果')
        if ($verifyErrors.Count -gt 0) {
            throw ('输出复核失败：' + ($verifyErrors -join '；'))
        }
        Update-OrganizerProgress -Message ($ProgressPrefix + '：正在完成输出文件夹')
        Move-Item -LiteralPath $temporaryPath -Destination $Audit.OutputPath
        return $Audit.OutputPath
    }
    catch {
        if (Test-Path -LiteralPath $temporaryPath -PathType Container) {
            $resolvedTemporary = [IO.Path]::GetFullPath($temporaryPath)
            if ($resolvedTemporary.StartsWith($expectedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
                Remove-Item -LiteralPath $resolvedTemporary -Recurse -Force
            }
        }
        throw
    }
}

function Show-OrganizerWindow {
    param(
        [string]$LibraryRoot,
        [object[]]$Candidates,
        [switch]$SmokeTest
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName Microsoft.VisualBasic

    $form = New-Object System.Windows.Forms.Form
    $form.Text = '本地漫画整理器'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(1600, 820)
    $form.MinimumSize = New-Object System.Drawing.Size(1420, 700)
    $form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = '合并、拆分与标准化漫画章节'
    $title.AutoSize = $true
    $title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 15, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(16, 14)
    $form.Controls.Add($title)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = '载入时优先采用元数据章节顺序；无有效配置时按名称排序。勾选“并入上一话”可把多个来源接成同一话。'
    $hint.AutoSize = $true
    $hint.ForeColor = [System.Drawing.Color]::DimGray
    $hint.Location = New-Object System.Drawing.Point(18, 48)
    $form.Controls.Add($hint)

    $sourceLabel = New-Object System.Windows.Forms.Label
    $sourceLabel.Text = '1. 选择来源漫画文件夹'
    $sourceLabel.AutoSize = $true
    $sourceLabel.Location = New-Object System.Drawing.Point(18, 82)
    $form.Controls.Add($sourceLabel)

    $sourceList = New-Object System.Windows.Forms.CheckedListBox
    $sourceList.Location = New-Object System.Drawing.Point(18, 106)
    $sourceList.Size = New-Object System.Drawing.Size(365, 438)
    $sourceList.Anchor = 'Top,Bottom,Left'
    $sourceList.CheckOnClick = $true
    $sourceList.HorizontalScrollbar = $true
    foreach ($candidate in $Candidates) {
        [void]$sourceList.Items.Add($candidate.Name)
    }
    $form.Controls.Add($sourceList)

    $loadSelected = New-Object System.Windows.Forms.Button
    $loadSelected.Text = '载入所选文件夹'
    $loadSelected.Location = New-Object System.Drawing.Point(18, 558)
    $loadSelected.Size = New-Object System.Drawing.Size(177, 36)
    $loadSelected.Anchor = 'Bottom,Left'
    $form.Controls.Add($loadSelected)

    $loadNewSources = New-Object System.Windows.Forms.Button
    $loadNewSources.Text = '载入新增文件夹'
    $loadNewSources.Location = New-Object System.Drawing.Point(203, 558)
    $loadNewSources.Size = New-Object System.Drawing.Size(180, 36)
    $loadNewSources.Anchor = 'Bottom,Left'
    $loadNewSources.Enabled = $false
    $form.Controls.Add($loadNewSources)

    $rescanSources = New-Object System.Windows.Forms.Button
    $rescanSources.Text = '重新扫描来源文件夹库'
    $rescanSources.Location = New-Object System.Drawing.Point(18, 600)
    $rescanSources.Size = New-Object System.Drawing.Size(365, 36)
    $rescanSources.Anchor = 'Bottom,Left'
    $form.Controls.Add($rescanSources)

    $outputLabel = New-Object System.Windows.Forms.Label
    $outputLabel.Text = '输出漫画名称：'
    $outputLabel.AutoSize = $true
    $outputLabel.Location = New-Object System.Drawing.Point(405, 82)
    $form.Controls.Add($outputLabel)

    $outputName = New-Object System.Windows.Forms.ComboBox
    $outputName.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $outputName.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
    $outputName.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
    $outputName.Location = New-Object System.Drawing.Point(510, 78)
    $outputName.Size = New-Object System.Drawing.Size(850, 27)
    $outputName.Anchor = 'Top,Left,Right'
    $form.Controls.Add($outputName)

    $descriptionButton = New-Object System.Windows.Forms.Button
    $descriptionButton.Text = '编辑简介（未填写）'
    $descriptionButton.Location = New-Object System.Drawing.Point(1370, 75)
    $descriptionButton.Size = New-Object System.Drawing.Size(195, 32)
    $descriptionButton.Anchor = 'Top,Right'
    $descriptionButton.BackColor = [System.Drawing.Color]::FromArgb(238, 244, 250)
    $form.Controls.Add($descriptionButton)

    $wholeCoverButton = New-Object System.Windows.Forms.Button
    $wholeCoverButton.Text = '整本封面：按来源'
    $wholeCoverButton.Location = New-Object System.Drawing.Point(405, 114)
    $wholeCoverButton.Size = New-Object System.Drawing.Size(190, 32)
    $wholeCoverButton.Anchor = 'Top,Left'
    $wholeCoverButton.BackColor = [System.Drawing.Color]::FromArgb(238, 244, 250)
    $form.Controls.Add($wholeCoverButton)

    $showChapterCovers = New-Object System.Windows.Forms.CheckBox
    $showChapterCovers.Text = '漫画目录显示每话封面（整本总开关）'
    $showChapterCovers.AutoSize = $true
    $showChapterCovers.Location = New-Object System.Drawing.Point(610, 120)
    $showChapterCovers.Anchor = 'Top,Left'
    $form.Controls.Add($showChapterCovers)

    $chapterCoverButton = New-Object System.Windows.Forms.Button
    $chapterCoverButton.Text = '设置选中话封面…'
    $chapterCoverButton.Location = New-Object System.Drawing.Point(1375, 114)
    $chapterCoverButton.Size = New-Object System.Drawing.Size(190, 32)
    $chapterCoverButton.Anchor = 'Top,Right'
    $form.Controls.Add($chapterCoverButton)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(405, 154)
    $grid.Size = New-Object System.Drawing.Size(1160, 390)
    $grid.Anchor = 'Top,Bottom,Left,Right'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.MultiSelect = $true
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.RowHeadersVisible = $false
    $grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    [void]$grid.Columns.Add('Number', '话序 / 特殊标签')
    [void]$grid.Columns.Add('Title', '本话章节名（可单独填写）')
    [void]$grid.Columns.Add('SourceFolder', '来源漫画文件夹')
    [void]$grid.Columns.Add('SourceChapter', '来源章节')
    [void]$grid.Columns.Add('Start', '起始')
    [void]$grid.Columns.Add('End', '结束')
    [void]$grid.Columns.Add('Total', '总数')
    $chapterCoverColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $chapterCoverColumn.Name = 'ChapterCover'
    $chapterCoverColumn.HeaderText = '章节封面 ▼'
    $chapterCoverColumn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    foreach ($coverChoice in @('跟随元数据', '首图', '自选', '隐藏')) { [void]$chapterCoverColumn.Items.Add($coverChoice) }
    [void]$grid.Columns.Add($chapterCoverColumn)
    [void]$grid.Columns.Add('ChapterCoverMode', '章节封面模式')
    [void]$grid.Columns.Add('ChapterCoverPath', '章节封面路径')
    $mergeColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $mergeColumn.Name = 'Merge'
    $mergeColumn.HeaderText = '并入上一话'
    $mergeColumn.ToolTipText = '勾选后，本行图片会接在上一行所属章节末尾；可跨来源文件夹合并为同一话'
    $mergeColumn.TrueValue = $true
    $mergeColumn.FalseValue = $false
    $mergeColumn.IndeterminateValue = $false
    $mergeColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    [void]$grid.Columns.Add($mergeColumn)
    $coverColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $coverColumn.Name = 'Cover'
    $coverColumn.HeaderText = '封面'
    $coverColumn.ToolTipText = '勾选这一行的来源漫画文件夹作为输出封面；没有 cover.jpg 时自动取该漫画第一章的 0001；只能选择一行'
    $coverColumn.TrueValue = $true
    $coverColumn.FalseValue = $false
    $coverColumn.IndeterminateValue = $false
    $coverColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    [void]$grid.Columns.Add($coverColumn)
    $batchColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $batchColumn.Name = 'Batch'
    $batchColumn.HeaderText = '选择'
    $batchColumn.ToolTipText = '勾选后，章节封面、合并、复制、删除、上移和下移等按钮会同时处理这些行'
    $batchColumn.TrueValue = $true
    $batchColumn.FalseValue = $false
    $batchColumn.IndeterminateValue = $false
    $batchColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    [void]$grid.Columns.Add($batchColumn)
    [void]$grid.Columns.Add('MetadataCoverInfo', '元数据实际指向')
    $grid.Columns['Batch'].DisplayIndex = 0
    $grid.Columns['Number'].Width = 85
    $grid.Columns['Number'].ToolTipText = '可填 4.5、1.01，也可直接填“特典话”“番外篇”“插画集”；实际阅读位置由行顺序决定'
    $grid.Columns['Title'].Width = 160
    $grid.Columns['Title'].ToolTipText = '每一行都可单独命名；输出示例：第4.5话 特典篇'
    $grid.Columns['SourceFolder'].Width = 150
    $grid.Columns['SourceChapter'].Width = 90
    $grid.Columns['Start'].Width = 65
    $grid.Columns['End'].Width = 65
    $grid.Columns['Total'].Width = 55
    $grid.Columns['ChapterCover'].Width = 115
    $grid.Columns['ChapterCover'].ToolTipText = '可直接下拉选择：首图、自选本地图片、隐藏或跟随来源元数据'
    $grid.Columns['ChapterCoverMode'].Visible = $false
    $grid.Columns['ChapterCoverPath'].Visible = $false
    $grid.Columns['Merge'].Width = 82
    $grid.Columns['Cover'].Width = 55
    $grid.Columns['Batch'].Width = 48
    $grid.Columns['MetadataCoverInfo'].Width = 170
    $grid.Columns['MetadataCoverInfo'].ReadOnly = $true
    $grid.Columns['MetadataCoverInfo'].ToolTipText = '当章节封面选择“跟随元数据”时，这里显示来源元数据最终会采用首图、隐藏还是资源目录中的自选图片'
    $grid.Columns['MetadataCoverInfo'].DisplayIndex = $grid.Columns['ChapterCover'].DisplayIndex + 1
    $grid.Columns['Total'].ReadOnly = $true
    $form.Controls.Add($grid)

    $splitRow = New-Object System.Windows.Forms.Button
    $splitRow.Text = '按范围拆分选中行'
    $splitRow.Location = New-Object System.Drawing.Point(405, 572)
    $splitRow.Size = New-Object System.Drawing.Size(155, 34)
    $splitRow.Anchor = 'Bottom,Left'
    $form.Controls.Add($splitRow)

    $duplicateRow = New-Object System.Windows.Forms.Button
    $duplicateRow.Text = '复制选中行'
    $duplicateRow.Location = New-Object System.Drawing.Point(568, 572)
    $duplicateRow.Size = New-Object System.Drawing.Size(115, 34)
    $duplicateRow.Anchor = 'Bottom,Left'
    $form.Controls.Add($duplicateRow)

    $deleteRow = New-Object System.Windows.Forms.Button
    $deleteRow.Text = '删除选中行'
    $deleteRow.Location = New-Object System.Drawing.Point(691, 572)
    $deleteRow.Size = New-Object System.Drawing.Size(115, 34)
    $deleteRow.Anchor = 'Bottom,Left'
    $form.Controls.Add($deleteRow)

    $mergeSelected = New-Object System.Windows.Forms.Button
    $mergeSelected.Text = '合并所选为同一话'
    $mergeSelected.Location = New-Object System.Drawing.Point(814, 572)
    $mergeSelected.Size = New-Object System.Drawing.Size(148, 34)
    $mergeSelected.Anchor = 'Bottom,Left'
    $form.Controls.Add($mergeSelected)

    $moveTop = New-Object System.Windows.Forms.Button
    $moveTop.Text = '移到最上'
    $moveTop.Location = New-Object System.Drawing.Point(970, 572)
    $moveTop.Size = New-Object System.Drawing.Size(88, 34)
    $moveTop.Anchor = 'Bottom,Left'
    $form.Controls.Add($moveTop)

    $moveUp = New-Object System.Windows.Forms.Button
    $moveUp.Text = '上移'
    $moveUp.Location = New-Object System.Drawing.Point(1066, 572)
    $moveUp.Size = New-Object System.Drawing.Size(70, 34)
    $moveUp.Anchor = 'Bottom,Left'
    $form.Controls.Add($moveUp)

    $moveDown = New-Object System.Windows.Forms.Button
    $moveDown.Text = '下移'
    $moveDown.Location = New-Object System.Drawing.Point(1144, 572)
    $moveDown.Size = New-Object System.Drawing.Size(70, 34)
    $moveDown.Anchor = 'Bottom,Left'
    $form.Controls.Add($moveDown)

    $moveBottom = New-Object System.Windows.Forms.Button
    $moveBottom.Text = '移到最下'
    $moveBottom.Location = New-Object System.Drawing.Point(1222, 572)
    $moveBottom.Size = New-Object System.Drawing.Size(88, 34)
    $moveBottom.Anchor = 'Bottom,Left'
    $form.Controls.Add($moveBottom)

    $autoNumber = New-Object System.Windows.Forms.Button
    $autoNumber.Text = '从选中行后续编号'
    $autoNumber.Location = New-Object System.Drawing.Point(1318, 572)
    $autoNumber.Size = New-Object System.Drawing.Size(155, 34)
    $autoNumber.Anchor = 'Bottom,Left'
    $form.Controls.Add($autoNumber)

    $selectAllRows = New-Object System.Windows.Forms.Button
    $selectAllRows.Text = '选择全部行'
    $selectAllRows.Location = New-Object System.Drawing.Point(405, 612)
    $selectAllRows.Size = New-Object System.Drawing.Size(110, 32)
    $selectAllRows.Anchor = 'Bottom,Left'
    $form.Controls.Add($selectAllRows)

    $clearSelectedRows = New-Object System.Windows.Forms.Button
    $clearSelectedRows.Text = '清空行选择'
    $clearSelectedRows.Location = New-Object System.Drawing.Point(523, 612)
    $clearSelectedRows.Size = New-Object System.Drawing.Size(110, 32)
    $clearSelectedRows.Anchor = 'Bottom,Left'
    $form.Controls.Add($clearSelectedRows)

    $validateButton = New-Object System.Windows.Forms.Button
    $validateButton.Text = '核验方案'
    $validateButton.Location = New-Object System.Drawing.Point(641, 612)
    $validateButton.Size = New-Object System.Drawing.Size(110, 38)
    $validateButton.Anchor = 'Bottom,Left'
    $form.Controls.Add($validateButton)

    $savePlan = New-Object System.Windows.Forms.Button
    $savePlan.Text = '保存方案'
    $savePlan.Location = New-Object System.Drawing.Point(759, 612)
    $savePlan.Size = New-Object System.Drawing.Size(110, 38)
    $savePlan.Anchor = 'Bottom,Left'
    $form.Controls.Add($savePlan)

    $loadPlan = New-Object System.Windows.Forms.Button
    $loadPlan.Text = '加载方案'
    $loadPlan.Location = New-Object System.Drawing.Point(877, 612)
    $loadPlan.Size = New-Object System.Drawing.Size(110, 38)
    $loadPlan.Anchor = 'Bottom,Left'
    $form.Controls.Add($loadPlan)

    $helpButton = New-Object System.Windows.Forms.Button
    $helpButton.Text = '使用说明'
    $helpButton.Location = New-Object System.Drawing.Point(995, 612)
    $helpButton.Size = New-Object System.Drawing.Size(110, 38)
    $helpButton.Anchor = 'Bottom,Left'
    $form.Controls.Add($helpButton)

    $clearWorkspace = New-Object System.Windows.Forms.Button
    $clearWorkspace.Text = '清空右侧内容'
    $clearWorkspace.Location = New-Object System.Drawing.Point(1113, 612)
    $clearWorkspace.Size = New-Object System.Drawing.Size(130, 38)
    $clearWorkspace.Anchor = 'Bottom,Left'
    $form.Controls.Add($clearWorkspace)

    $undoButton = New-Object System.Windows.Forms.Button
    $undoButton.Text = '撤销'
    $undoButton.Location = New-Object System.Drawing.Point(1251, 612)
    $undoButton.Size = New-Object System.Drawing.Size(80, 38)
    $undoButton.Anchor = 'Bottom,Left'
    $undoButton.Enabled = $false
    $form.Controls.Add($undoButton)

    $redoButton = New-Object System.Windows.Forms.Button
    $redoButton.Text = '恢复'
    $redoButton.Location = New-Object System.Drawing.Point(1339, 612)
    $redoButton.Size = New-Object System.Drawing.Size(80, 38)
    $redoButton.Anchor = 'Bottom,Left'
    $redoButton.Enabled = $false
    $form.Controls.Add($redoButton)

    $generateButton = New-Object System.Windows.Forms.Button
    $generateButton.Text = '开始整理'
    $generateButton.Location = New-Object System.Drawing.Point(1425, 612)
    $generateButton.Size = New-Object System.Drawing.Size(140, 38)
    $generateButton.Anchor = 'Bottom,Right'
    $generateButton.BackColor = [System.Drawing.Color]::FromArgb(35, 105, 160)
    $generateButton.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($generateButton)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = '等待载入来源文件夹。'
    $status.AutoSize = $false
    $status.Location = New-Object System.Drawing.Point(18, 666)
    $status.Size = New-Object System.Drawing.Size(1547, 70)
    $status.Anchor = 'Bottom,Left,Right'
    $status.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($status)

    $script:lastDefaultOutput = ''
    $script:OrganizerDescription = ''
    $script:OrganizerDescriptionSource = ''
    $script:OrganizerDescriptionLocked = $false
    $script:OrganizerLoadedOnce = $false
    $script:OrganizerCustomCoverPath = ''
    $historyState = [pscustomobject]@{
        Undo = New-Object System.Collections.ArrayList
        Redo = New-Object System.Collections.ArrayList
        Restoring = $false
        EditSnapshotTaken = $false
        Limit = 40
    }

    $operationControls = @(
        $sourceList, $loadSelected, $loadNewSources, $rescanSources, $outputName, $descriptionButton, $wholeCoverButton, $showChapterCovers, $chapterCoverButton,
        $grid, $splitRow, $duplicateRow, $deleteRow, $mergeSelected, $moveTop, $moveUp, $moveDown, $moveBottom, $autoNumber, $selectAllRows, $clearSelectedRows,
        $validateButton, $savePlan, $loadPlan, $helpButton, $clearWorkspace, $undoButton, $redoButton, $generateButton
    )
    $setOrganizerBusy = {
        param([bool]$Busy)
        if ($Busy) {
            $script:OrganizerBusyControlStates = @($operationControls | ForEach-Object {
                [pscustomobject]@{ Control = $_; Enabled = $_.Enabled }
            })
            foreach ($control in $operationControls) { $control.Enabled = $false }
        }
        else {
            foreach ($state in @($script:OrganizerBusyControlStates)) {
                if ($null -ne $state.Control -and -not $state.Control.IsDisposed) { $state.Control.Enabled = [bool]$state.Enabled }
            }
            $script:OrganizerBusyControlStates = @()
        }
        $script:OrganizerIsRunning = $Busy
        $form.UseWaitCursor = $Busy
        $form.Cursor = if ($Busy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
    $script:OrganizerProgressCallback = {
        param([string]$Message)
        if (-not $status.IsDisposed) {
            $status.Text = $Message
            $status.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    $showMessage = {
        param([string]$Message, [string]$Caption, [System.Windows.Forms.MessageBoxIcon]$Icon)
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $Message,
            $Caption,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $Icon
        ) | Out-Null
    }

    $getMetadataCoverDisplay = {
        param([AllowNull()][object]$SourceInfo, [string]$ComicPath)
        if ($null -eq $SourceInfo) {
            return [pscustomobject]@{ Display = '无法读取'; Details = '无法读取来源章节的元数据封面设置。' }
        }
        $hasConfiguredCover = if ($null -ne $SourceInfo.PSObject.Properties['HasConfiguredCover']) {
            [bool]$SourceInfo.HasConfiguredCover
        }
        else { $null -ne $SourceInfo.PSObject.Properties['ConfiguredCoverMode'] }
        if (-not $hasConfiguredCover) {
            return [pscustomobject]@{ Display = '无设置 → 首图'; Details = '来源没有逐话章节封面元数据；跟随元数据时回退到本话首图。' }
        }
        $showCovers = $null -ne $SourceInfo.PSObject.Properties['ConfiguredShowChapterCovers'] -and [bool]$SourceInfo.ConfiguredShowChapterCovers
        if (-not $showCovers) {
            return [pscustomobject]@{ Display = '总开关关闭 → 隐藏'; Details = '来源元数据 readerOptions.showChapterCovers 为 false；跟随元数据时不会显示章节封面。' }
        }
        $mode = ([string]$SourceInfo.ConfiguredCoverMode).Trim().ToLowerInvariant()
        if ($mode -eq 'none') {
            return [pscustomobject]@{ Display = '隐藏'; Details = '来源元数据把这一话的 coverMode 设为 none（隐藏）。' }
        }
        if ($mode -eq 'custom') {
            $relativePath = ([string]$SourceInfo.ConfiguredCoverFile).Trim()
            $displayName = if ([string]::IsNullOrWhiteSpace($relativePath)) { '路径缺失' } else { [IO.Path]::GetFileName($relativePath) }
            $resolvedPath = if ([string]::IsNullOrWhiteSpace($ComicPath)) { '' } else { Resolve-ConfiguredSourceCoverPath -ComicPath $ComicPath -RelativePath $relativePath }
            $details = '来源元数据自选封面：' + $relativePath
            if (-not [string]::IsNullOrWhiteSpace($resolvedPath)) { $details += "`r`n本地位置：" + $resolvedPath }
            else { $details += "`r`n注意：当前路径无法解析或图片已不存在。" }
            return [pscustomobject]@{ Display = '自选：' + $displayName; Details = $details }
        }
        return [pscustomobject]@{ Display = '首图'; Details = '来源元数据把这一话设为首图；跟随元数据时采用该话第一张正文图片。' }
    }

    $setMetadataCoverCell = {
        param([System.Windows.Forms.DataGridViewRow]$Row, [AllowNull()][object]$SourceInfo, [string]$ComicPath)
        $coverDisplay = & $getMetadataCoverDisplay $SourceInfo $ComicPath
        $Row.Cells['MetadataCoverInfo'].Value = [string]$coverDisplay.Display
        $Row.Cells['MetadataCoverInfo'].ToolTipText = [string]$coverDisplay.Details
        $Row.Cells['ChapterCover'].ToolTipText = '跟随元数据时：' + [string]$coverDisplay.Details
    }

    $updateDescriptionButton = {
        $length = $script:OrganizerDescription.Length
        if (-not $script:OrganizerDescriptionLocked -and -not [string]::IsNullOrWhiteSpace($script:OrganizerDescriptionSource)) {
            $descriptionButton.Text = ('简介：随漫画名（{0}字）' -f $length)
        }
        elseif ($script:OrganizerDescriptionLocked) {
            $descriptionButton.Text = if ($length -eq 0) { '编辑简介（已固定为空）' } else { '编辑简介（已固定 {0}字）' -f $length }
        }
        elseif ($length -eq 0) {
            $descriptionButton.Text = '编辑简介（未填写）'
        }
        else {
            $descriptionButton.Text = ('编辑简介（{0}字）' -f $length)
        }
    }

    $syncDefaultDescriptionFromOutputName = {
        if ($script:OrganizerDescriptionLocked -or $outputName.SelectedIndex -lt 0) { return }
        $sourceName = [string]$outputName.SelectedItem
        if ([string]::IsNullOrWhiteSpace($sourceName)) { return }
        $script:OrganizerDescriptionSource = $sourceName
        $script:OrganizerDescription = Get-SourceDescription -LibraryRoot $LibraryRoot -SourceFolder $sourceName
        & $updateDescriptionButton
    }

    $refreshOutputNameChoices = {
        param([string[]]$PreferredSources, [string]$PreferredText)
        $sourceNames = @($PreferredSources | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        $outputName.BeginUpdate()
        try {
            $outputName.Items.Clear()
            foreach ($sourceName in $sourceNames) { [void]$outputName.Items.Add($sourceName) }
        }
        finally { $outputName.EndUpdate() }
        if (-not [string]::IsNullOrWhiteSpace($PreferredText)) {
            $matchingIndex = $outputName.Items.IndexOf($PreferredText)
            if ($matchingIndex -ge 0) { $outputName.SelectedIndex = $matchingIndex } else { $outputName.Text = $PreferredText }
        }
        elseif ($outputName.Items.Count -gt 0) {
            $outputName.SelectedIndex = 0
        }
    }

    $outputName.add_SelectedIndexChanged({ if (-not $historyState.Restoring) { & $syncDefaultDescriptionFromOutputName } })
    $outputName.add_Enter({ if (-not $historyState.Restoring) { & $pushUndoSnapshot } })
    $showChapterCovers.add_MouseDown({ if (-not $historyState.Restoring) { & $pushUndoSnapshot } })

    $descriptionButton.add_Click({
        $sourceNames = @($sourceList.CheckedItems | ForEach-Object { [string]$_ })
        if ($sourceNames.Count -eq 0) {
            $sourceNames = @($grid.Rows | ForEach-Object { [string]$_.Cells['SourceFolder'].Value } | Where-Object { $_ } | Select-Object -Unique)
        }
        $descriptions = @{}
        foreach ($sourceName in $sourceNames) {
            $sourceDescription = Get-SourceDescription -LibraryRoot $LibraryRoot -SourceFolder $sourceName
            if (-not [string]::IsNullOrWhiteSpace($sourceDescription)) {
                $descriptions[$sourceName] = $sourceDescription
            }
        }

        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = '选择并编辑漫画简介'
        $dialog.StartPosition = 'CenterParent'
        $dialog.Size = New-Object System.Drawing.Size(820, 610)
        $dialog.MinimumSize = New-Object System.Drawing.Size(680, 500)
        $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
        $dialog.MinimizeBox = $false

        $sourceCaption = New-Object System.Windows.Forms.Label
        $sourceCaption.Text = '简介来源（只列出含简介的已导入文件夹）：'
        $sourceCaption.AutoSize = $true
        $sourceCaption.Location = New-Object System.Drawing.Point(18, 18)
        $dialog.Controls.Add($sourceCaption)

        $sourceCombo = New-Object System.Windows.Forms.ComboBox
        $sourceCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $sourceCombo.Location = New-Object System.Drawing.Point(18, 48)
        $sourceCombo.Size = New-Object System.Drawing.Size(620, 30)
        $sourceCombo.Anchor = 'Top,Left,Right'
        [void]$sourceCombo.Items.Add('手动编辑 / 留空')
        foreach ($sourceName in @($descriptions.Keys | Sort-Object)) { [void]$sourceCombo.Items.Add($sourceName) }
        $sourceCombo.SelectedIndex = 0
        if (-not [string]::IsNullOrWhiteSpace($script:OrganizerDescriptionSource)) {
            $existingIndex = $sourceCombo.Items.IndexOf($script:OrganizerDescriptionSource)
            if ($existingIndex -ge 0) { $sourceCombo.SelectedIndex = $existingIndex }
        }
        $dialog.Controls.Add($sourceCombo)

        $adoptButton = New-Object System.Windows.Forms.Button
        $adoptButton.Text = '采用所选简介'
        $adoptButton.Location = New-Object System.Drawing.Point(648, 46)
        $adoptButton.Size = New-Object System.Drawing.Size(135, 34)
        $adoptButton.Anchor = 'Top,Right'
        $dialog.Controls.Add($adoptButton)

        $editCaption = New-Object System.Windows.Forms.Label
        $editCaption.Text = '最终简介（可继续修改；清空后输出空简介）：'
        $editCaption.AutoSize = $true
        $editCaption.Location = New-Object System.Drawing.Point(18, 96)
        $dialog.Controls.Add($editCaption)

        $editor = New-Object System.Windows.Forms.TextBox
        $editor.Multiline = $true
        $editor.AcceptsReturn = $true
        $editor.ScrollBars = 'Vertical'
        $editor.WordWrap = $true
        $editor.Location = New-Object System.Drawing.Point(18, 126)
        $editor.Size = New-Object System.Drawing.Size(765, 365)
        $editor.Anchor = 'Top,Bottom,Left,Right'
        $editor.Text = $script:OrganizerDescription
        $dialog.Controls.Add($editor)

        $countLabel = New-Object System.Windows.Forms.Label
        $countLabel.AutoSize = $true
        $countLabel.Location = New-Object System.Drawing.Point(18, 505)
        $countLabel.Anchor = 'Bottom,Left'
        $dialog.Controls.Add($countLabel)
        $refreshCount = { $countLabel.Text = ('当前 {0} 字' -f $editor.Text.Length) }
        $editor.add_TextChanged({ & $refreshCount })
        & $refreshCount

        $adoptButton.add_Click({
            if ($sourceCombo.SelectedIndex -le 0) {
                $editor.Clear()
                return
            }
            $selectedSource = [string]$sourceCombo.SelectedItem
            $editor.Text = [string]$descriptions[$selectedSource]
        })

        $okDescription = New-Object System.Windows.Forms.Button
        $okDescription.Text = '确定'
        $okDescription.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $okDescription.Location = New-Object System.Drawing.Point(574, 520)
        $okDescription.Size = New-Object System.Drawing.Size(100, 36)
        $okDescription.Anchor = 'Bottom,Right'
        $dialog.Controls.Add($okDescription)

        $cancelDescription = New-Object System.Windows.Forms.Button
        $cancelDescription.Text = '取消'
        $cancelDescription.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $cancelDescription.Location = New-Object System.Drawing.Point(683, 520)
        $cancelDescription.Size = New-Object System.Drawing.Size(100, 36)
        $cancelDescription.Anchor = 'Bottom,Right'
        $dialog.Controls.Add($cancelDescription)
        $dialog.AcceptButton = $okDescription
        $dialog.CancelButton = $cancelDescription

        if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            & $pushUndoSnapshot
            $script:OrganizerDescription = $editor.Text.Trim()
            $script:OrganizerDescriptionSource = if ($sourceCombo.SelectedIndex -gt 0) { [string]$sourceCombo.SelectedItem } else { '' }
            $script:OrganizerDescriptionLocked = $true
            & $updateDescriptionButton
            $status.Text = if ($script:OrganizerDescription.Length -gt 0) {
                ('已设置漫画简介：{0} 字；最终仍可再次编辑。' -f $script:OrganizerDescription.Length)
            }
            else { '已将输出漫画简介设为空。' }
        }
        $dialog.Dispose()
    })

    $updateWholeCoverButton = {
        $wholeCoverButton.Text = if ([string]::IsNullOrWhiteSpace($script:OrganizerCustomCoverPath)) {
            '整本封面：按来源'
        }
        else {
            '整本封面：自选 ' + [IO.Path]::GetFileName($script:OrganizerCustomCoverPath)
        }
    }
    $wholeCoverMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $chooseWholeCover = $wholeCoverMenu.Items.Add('自选任意本地图片…')
    $restoreWholeCover = $wholeCoverMenu.Items.Add('恢复按表格“封面”来源')
    $chooseWholeCover.add_Click({
        $picker = New-Object System.Windows.Forms.OpenFileDialog
        $picker.Title = '选择输出漫画封面'
        $picker.Filter = '图片文件|*.jpg;*.jpeg;*.png;*.webp;*.gif;*.bmp;*.avif|所有文件|*.*'
        if ($picker.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            & $pushUndoSnapshot
            $script:OrganizerCustomCoverPath = $picker.FileName
            & $updateWholeCoverButton
            $status.Text = '已选择自定义整本封面；整理时会复制为输出漫画的 cover 文件。'
        }
    })
    $restoreWholeCover.add_Click({
        & $pushUndoSnapshot
        $script:OrganizerCustomCoverPath = ''
        & $updateWholeCoverButton
        $status.Text = '整本封面已恢复为表格中勾选的来源漫画封面。'
    })
    $wholeCoverButton.add_Click({ $wholeCoverMenu.Show($wholeCoverButton, 0, $wholeCoverButton.Height) })
    & $updateWholeCoverButton

    $chapterCoverUiState = [pscustomobject]@{ Changing = $false; Picking = $false }
    $mergedDisplayStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $mergedDisplayStyle.ForeColor = [System.Drawing.Color]::Gray
    $mergedDisplayStyle.BackColor = [System.Drawing.Color]::FromArgb(242, 242, 242)
    $mergedDisplayStyle.SelectionForeColor = [System.Drawing.Color]::DimGray
    $mergedDisplayStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $applyMergedRowDisplay = {
        param([System.Windows.Forms.DataGridViewRow]$Row)
        if ($null -eq $Row) { return }
        $isMerged = ($Row.Cells['Merge'].Value -eq $true)
        foreach ($columnName in @('Number', 'Title', 'ChapterCover', 'MetadataCoverInfo')) {
            $cell = $Row.Cells[$columnName]
            $cell.Style = if ($isMerged) { $mergedDisplayStyle } else { New-Object System.Windows.Forms.DataGridViewCellStyle }
        }
        $Row.Cells['Number'].ReadOnly = $isMerged
        $Row.Cells['Title'].ReadOnly = $isMerged
        $Row.Cells['ChapterCover'].ReadOnly = $isMerged
        $Row.Cells['MetadataCoverInfo'].ReadOnly = $true
    }
    $captureOrganizerState = {
        $rowStates = @()
        foreach ($row in $grid.Rows) {
            $values = @()
            for ($columnIndex = 0; $columnIndex -lt $grid.Columns.Count; $columnIndex++) { $values += ,$row.Cells[$columnIndex].Value }
            $rowStates += [pscustomobject]@{
                Values = $values
                MetadataToolTip = [string]$row.Cells['MetadataCoverInfo'].ToolTipText
                ChapterCoverToolTip = [string]$row.Cells['ChapterCover'].ToolTipText
            }
        }
        $sourceStates = @()
        for ($sourceIndex = 0; $sourceIndex -lt $sourceList.Items.Count; $sourceIndex++) {
            $sourceStates += [pscustomobject]@{ Name = [string]$sourceList.Items[$sourceIndex]; Checked = $sourceList.GetItemChecked($sourceIndex) }
        }
        return [pscustomobject]@{
            Rows = $rowStates
            Sources = $sourceStates
            OutputChoices = @($outputName.Items | ForEach-Object { [string]$_ })
            OutputName = [string]$outputName.Text
            Description = [string]$script:OrganizerDescription
            DescriptionSource = [string]$script:OrganizerDescriptionSource
            DescriptionLocked = [bool]$script:OrganizerDescriptionLocked
            CustomCoverPath = [string]$script:OrganizerCustomCoverPath
            ShowChapterCovers = [bool]$showChapterCovers.Checked
            LastDefaultOutput = [string]$script:lastDefaultOutput
            LoadedOnce = [bool]$script:OrganizerLoadedOnce
            LoadSelectedText = [string]$loadSelected.Text
            LoadNewEnabled = [bool]$loadNewSources.Enabled
            CurrentRow = if ($null -ne $grid.CurrentRow) { [int]$grid.CurrentRow.Index } else { -1 }
            CurrentColumn = if ($null -ne $grid.CurrentCell) { [string]$grid.CurrentCell.OwningColumn.Name } else { 'Number' }
        }
    }
    $restoreOrganizerState = {
        param([object]$Snapshot)
        $historyState.Restoring = $true
        $chapterCoverUiState.Changing = $true
        try {
            $sourceList.BeginUpdate()
            try {
                $sourceList.Items.Clear()
                foreach ($sourceState in @($Snapshot.Sources)) {
                    $sourceIndex = $sourceList.Items.Add([string]$sourceState.Name)
                    $sourceList.SetItemChecked($sourceIndex, [bool]$sourceState.Checked)
                }
            }
            finally { $sourceList.EndUpdate() }

            $outputName.BeginUpdate()
            try {
                $outputName.Items.Clear()
                foreach ($choice in @($Snapshot.OutputChoices)) { [void]$outputName.Items.Add([string]$choice) }
            }
            finally { $outputName.EndUpdate() }
            $outputName.Text = [string]$Snapshot.OutputName

            $grid.SuspendLayout()
            try {
                $grid.Rows.Clear()
                foreach ($rowState in @($Snapshot.Rows)) {
                    $rowIndex = $grid.Rows.Add()
                    $row = $grid.Rows[$rowIndex]
                    for ($columnIndex = 0; $columnIndex -lt $grid.Columns.Count; $columnIndex++) { $row.Cells[$columnIndex].Value = $rowState.Values[$columnIndex] }
                    $row.Cells['MetadataCoverInfo'].ToolTipText = [string]$rowState.MetadataToolTip
                    $row.Cells['ChapterCover'].ToolTipText = [string]$rowState.ChapterCoverToolTip
                    & $applyMergedRowDisplay $row
                }
            }
            finally { $grid.ResumeLayout() }

            $script:OrganizerDescription = [string]$Snapshot.Description
            $script:OrganizerDescriptionSource = [string]$Snapshot.DescriptionSource
            $script:OrganizerDescriptionLocked = [bool]$Snapshot.DescriptionLocked
            $script:OrganizerCustomCoverPath = [string]$Snapshot.CustomCoverPath
            $showChapterCovers.Checked = [bool]$Snapshot.ShowChapterCovers
            $script:lastDefaultOutput = [string]$Snapshot.LastDefaultOutput
            $script:OrganizerLoadedOnce = [bool]$Snapshot.LoadedOnce
            $loadSelected.Text = [string]$Snapshot.LoadSelectedText
            $loadNewSources.Enabled = [bool]$Snapshot.LoadNewEnabled
            & $updateDescriptionButton
            & $updateWholeCoverButton
            if ($Snapshot.CurrentRow -ge 0 -and $Snapshot.CurrentRow -lt $grid.Rows.Count) {
                $column = if ($grid.Columns.Contains([string]$Snapshot.CurrentColumn)) { $grid.Columns[[string]$Snapshot.CurrentColumn] } else { $grid.Columns['Number'] }
                $grid.CurrentCell = $grid.Rows[[int]$Snapshot.CurrentRow].Cells[$column.Index]
            }
        }
        finally {
            $chapterCoverUiState.Changing = $false
            $historyState.Restoring = $false
        }
    }
    $updateHistoryButtons = {
        $undoButton.Enabled = $historyState.Undo.Count -gt 0
        $redoButton.Enabled = $historyState.Redo.Count -gt 0
    }
    $pushUndoSnapshot = {
        if ($historyState.Restoring) { return }
        [void]$historyState.Undo.Add((& $captureOrganizerState))
        while ($historyState.Undo.Count -gt $historyState.Limit) { $historyState.Undo.RemoveAt(0) }
        $historyState.Redo.Clear()
        & $updateHistoryButtons
    }
    $invokeUndo = {
        if ($historyState.Undo.Count -eq 0) { return }
        $current = & $captureOrganizerState
        $targetIndex = $historyState.Undo.Count - 1
        $target = $historyState.Undo[$targetIndex]
        $historyState.Undo.RemoveAt($targetIndex)
        [void]$historyState.Redo.Add($current)
        & $restoreOrganizerState $target
        & $updateHistoryButtons
        $status.Text = '已撤销上一个操作。'
    }
    $invokeRedo = {
        if ($historyState.Redo.Count -eq 0) { return }
        $current = & $captureOrganizerState
        $targetIndex = $historyState.Redo.Count - 1
        $target = $historyState.Redo[$targetIndex]
        $historyState.Redo.RemoveAt($targetIndex)
        [void]$historyState.Undo.Add($current)
        & $restoreOrganizerState $target
        & $updateHistoryButtons
        $status.Text = '已恢复上一个被撤销的操作。'
    }
    $undoButton.add_Click({ & $invokeUndo })
    $redoButton.add_Click({ & $invokeRedo })
    $form.KeyPreview = $true
    $form.add_KeyDown({
        param($sender, $eventArgs)
        if ($eventArgs.Control -and $eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Z) { & $invokeUndo; $eventArgs.SuppressKeyPress = $true }
        elseif ($eventArgs.Control -and ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Y -or ($eventArgs.Shift -and $eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Z))) { & $invokeRedo; $eventArgs.SuppressKeyPress = $true }
    })
    $getBatchRows = {
        $rows = @($grid.Rows | Where-Object { $_.Cells['Batch'].Value -eq $true })
        if ($rows.Count -eq 0 -and $null -ne $grid.CurrentRow) { $rows = @($grid.CurrentRow) }
        return @($rows)
    }
    $setChapterCoverRows = {
        param([object[]]$Rows, [string]$Mode, [string]$Path)
        $usableRows = @($Rows | Where-Object { $_.Cells['Merge'].Value -ne $true })
        if ($usableRows.Count -eq 0) {
            & $showMessage '所选行都会并入其他章节；请勾选该话的第一行设置封面。' '章节封面' ([System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        & $pushUndoSnapshot
        $display = switch ($Mode) { 'metadata' { '跟随元数据' } 'custom' { '自选' } 'none' { '隐藏' } default { '首图' } }
        $chapterCoverUiState.Changing = $true
        try {
            foreach ($row in $usableRows) {
                $row.Cells['ChapterCoverMode'].Value = $Mode
                $row.Cells['ChapterCoverPath'].Value = if ($Mode -eq 'custom') { $Path } else { '' }
                $row.Cells['ChapterCover'].Value = $display
            }
        }
        finally { $chapterCoverUiState.Changing = $false }
        if ($Mode -in @('first', 'custom')) { $showChapterCovers.Checked = $true }
        $skippedCount = @($Rows).Count - $usableRows.Count
        if ($Mode -eq 'metadata') {
            $actualValues = @($usableRows | ForEach-Object { [string]$_.Cells['MetadataCoverInfo'].Value } | Where-Object { $_ } | Select-Object -Unique)
            $status.Text = ('已设为“跟随元数据”；当前实际指向：{0}{1}。详细路径可查看右侧“元数据实际指向”列或悬停提示。' -f ($actualValues -join '、'), $(if ($skippedCount -gt 0) { '；跳过 ' + $skippedCount + ' 行并入内容' } else { '' }))
        }
        else {
            $status.Text = ('已把 {0} 行章节封面设为“{1}”{2}。' -f $usableRows.Count, $display, $(if ($skippedCount -gt 0) { '，并跳过 ' + $skippedCount + ' 行“并入上一话”' } else { '' }))
        }
    }
    $applyChapterCoverChoice = {
        param([string]$Mode)
        $rows = @(& $getBatchRows)
        if ($rows.Count -eq 0) {
            & $showMessage '请先勾选第一列中的一行或多行；未勾选时也可直接点中一行再操作。' '章节封面' ([System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $path = ''
        if ($Mode -eq 'custom') {
            $picker = New-Object System.Windows.Forms.OpenFileDialog
            $picker.Title = if ($rows.Count -gt 1) { '选择要统一用于所选章节的目录封面' } else { '选择这一话的目录封面' }
            $picker.Filter = '图片文件|*.jpg;*.jpeg;*.png;*.webp;*.gif;*.bmp;*.avif|所有文件|*.*'
            if ($picker.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $path = $picker.FileName
            if ($rows.Count -gt 1) {
                $answer = [System.Windows.Forms.MessageBox]::Show($form, ('将同一张图片设为所选 {0} 行的章节封面，是否继续？' -f $rows.Count), '批量设置自选封面', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }
        }
        & $setChapterCoverRows $rows $Mode $path
    }
    $chapterCoverMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $metadataChapterCover = $chapterCoverMenu.Items.Add('跟随元数据')
    $firstPageCover = $chapterCoverMenu.Items.Add('首图')
    $customChapterCover = $chapterCoverMenu.Items.Add('自选任意本地图片…')
    $noChapterCover = $chapterCoverMenu.Items.Add('隐藏')
    $metadataChapterCover.add_Click({ & $applyChapterCoverChoice 'metadata' })
    $firstPageCover.add_Click({ & $applyChapterCoverChoice 'first' })
    $customChapterCover.add_Click({ & $applyChapterCoverChoice 'custom' })
    $noChapterCover.add_Click({ & $applyChapterCoverChoice 'none' })
    $chapterCoverButton.add_Click({ $chapterCoverMenu.Show($chapterCoverButton, 0, $chapterCoverButton.Height) })

    $renumberRows = {
        if ($grid.Rows.Count -eq 0) { return }
        & $pushUndoSnapshot
        $logicalRows = @($grid.Rows | Where-Object { $_.Cells['Merge'].Value -ne $true } | Sort-Object Index)
        if ($logicalRows.Count -eq 0) {
            & $showMessage '当前表格没有独立章节起始行，请先取消至少一行的“并入上一话”。' '无法继续编号' ([System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $getLogicalOwnerRow = {
            param([System.Windows.Forms.DataGridViewRow]$InputRow)
            $ownerIndex = $InputRow.Index
            while ($ownerIndex -gt 0 -and $grid.Rows[$ownerIndex].Cells['Merge'].Value -eq $true) { $ownerIndex-- }
            return $grid.Rows[$ownerIndex]
        }
        $batchRows = @($grid.Rows | Where-Object { $_.Cells['Batch'].Value -eq $true } | Sort-Object Index)
        $selectedLogicalRows = @()
        $seenLogicalRows = @{}
        foreach ($batchRow in $batchRows) {
            $ownerRow = & $getLogicalOwnerRow $batchRow
            if (-not $seenLogicalRows.ContainsKey([string]$ownerRow.Index)) {
                $seenLogicalRows[[string]$ownerRow.Index] = $true
                $selectedLogicalRows += $ownerRow
            }
        }
        $selectedLogicalRows = @($selectedLogicalRows | Sort-Object Index)
        $anchorRow = $null
        $usedBatchRange = $false
        if ($selectedLogicalRows.Count -gt 1) {
            $isContinuous = $true
            for ($selectionIndex = 1; $selectionIndex -lt $selectedLogicalRows.Count; $selectionIndex++) {
                $previousLogicalIndex = [array]::IndexOf($logicalRows, $selectedLogicalRows[$selectionIndex - 1])
                $currentLogicalIndex = [array]::IndexOf($logicalRows, $selectedLogicalRows[$selectionIndex])
                if ($currentLogicalIndex -ne ($previousLogicalIndex + 1)) {
                    $isContinuous = $false
                    break
                }
            }
            if (-not $isContinuous) {
                & $showMessage '第一列勾选了多段不连续的行，无法判断从哪一段之后续编。请只保留一段连续勾选，或清空勾选后点中一行作为锚点。' '无法继续编号' ([System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }
            $anchorRow = $selectedLogicalRows[-1]
            $usedBatchRange = $true
        }
        elseif ($selectedLogicalRows.Count -eq 1) {
            $anchorRow = $selectedLogicalRows[0]
        }
        elseif ($null -ne $grid.CurrentRow) {
            $anchorRow = & $getLogicalOwnerRow $grid.CurrentRow
        }
        if ($null -eq $anchorRow) {
            $nextNumber = 1
            foreach ($row in $logicalRows) {
                $row.Cells['Number'].Value = $nextNumber
                $nextNumber++
            }
            $status.Text = ('没有选中锚点，已按 {0} 个合并后的逻辑章节从 1 开始编号；并入行保留原话序灰显，但不会参与编号。' -f $logicalRows.Count)
            return
        }
        $anchorLogicalIndex = [array]::IndexOf($logicalRows, $anchorRow)
        $anchorNumber = ConvertTo-ChapterNumberInfo -Value $anchorRow.Cells['Number'].Value
        if ($null -eq $anchorNumber) {
            & $showMessage '作为锚点的最后一行话序无效；请先填写正整数或小数，例如 3 或 4.5。' '无法继续编号' ([System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $nextNumber = [int][decimal]::Floor($anchorNumber.Value) + 1
        for ($logicalIndex = $anchorLogicalIndex + 1; $logicalIndex -lt $logicalRows.Count; $logicalIndex++) {
            $logicalRows[$logicalIndex].Cells['Number'].Value = $nextNumber
            $nextNumber++
        }
        $status.Text = if ($usedBatchRange) {
            '已采用连续勾选的最后一个逻辑章节作为锚点，只给合并后的独立章节续编；并入行原值保留灰显且不占号。'
        }
        else { '已从所选逻辑章节的下一话开始续编；并入行原话序保留灰显，但不再占用章节编号。' }
    }

    $coverSelectionState = [pscustomobject]@{ Changing = $false }
    $setCoverRow = {
        param([int]$RowIndex)
        if ($RowIndex -lt 0 -or $RowIndex -ge $grid.Rows.Count) { return }
        $coverSelectionState.Changing = $true
        try {
            foreach ($row in $grid.Rows) {
                $row.Cells['Cover'].Value = ($row.Index -eq $RowIndex)
            }
        }
        finally {
            $coverSelectionState.Changing = $false
        }
    }
    $grid.add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty -and $null -ne $grid.CurrentCell -and $grid.CurrentCell.OwningColumn.Name -in @('Cover', 'Batch', 'Merge', 'ChapterCover')) {
            [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })
    $grid.add_CellBeginEdit({
        param($sender, $eventArgs)
        if ($historyState.Restoring -or $eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -lt 0) { return }
        $columnName = $grid.Columns[$eventArgs.ColumnIndex].Name
        if ($columnName -in @('Number', 'Title', 'Start', 'End', 'Merge', 'Cover')) { & $pushUndoSnapshot }
    })
    $grid.add_DataError({ param($sender, $eventArgs); $eventArgs.ThrowException = $false })
    $grid.add_CellValueChanged({
        param($sender, $eventArgs)
        if ($historyState.Restoring -or $coverSelectionState.Changing -or $chapterCoverUiState.Changing -or $eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -lt 0) { return }
        $row = $grid.Rows[$eventArgs.RowIndex]
        $columnName = $grid.Columns[$eventArgs.ColumnIndex].Name
        if ($columnName -eq 'ChapterCover') {
            $choice = [string]$row.Cells['ChapterCover'].Value
            $mode = switch ($choice) { '跟随元数据' { 'metadata' } '自选' { 'custom' } '隐藏' { 'none' } default { 'first' } }
            if ($mode -eq 'custom' -and [string]::IsNullOrWhiteSpace([string]$row.Cells['ChapterCoverPath'].Value) -and -not $chapterCoverUiState.Picking) {
                $chapterCoverUiState.Picking = $true
                try {
                    $picker = New-Object System.Windows.Forms.OpenFileDialog
                    $picker.Title = '选择这一话的目录封面'
                    $picker.Filter = '图片文件|*.jpg;*.jpeg;*.png;*.webp;*.gif;*.bmp;*.avif|所有文件|*.*'
                    if ($picker.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                        & $setChapterCoverRows @($row) 'custom' $picker.FileName
                    }
                    else { & $setChapterCoverRows @($row) 'first' '' }
                }
                finally { $chapterCoverUiState.Picking = $false }
            }
            else {
                & $setChapterCoverRows @($row) $mode ([string]$row.Cells['ChapterCoverPath'].Value)
            }
            return
        }
        if ($columnName -eq 'Merge') {
            & $applyMergedRowDisplay $row
            $grid.InvalidateRow($eventArgs.RowIndex)
            $status.Text = if ($row.Cells['Merge'].Value -eq $true) {
                '本行已并入上一话；原话序、章节名和章节封面信息会保留并灰显，但输出时不单独生效。'
            }
            else { '本行已恢复为独立章节；原话序、章节名和章节封面信息重新生效。' }
            return
        }
        if ($columnName -ne 'Cover') { return }
        if ($row.Cells['Cover'].Value -eq $true) {
            & $setCoverRow $eventArgs.RowIndex
            return
        }
        $hasCover = $false
        foreach ($candidateRow in $grid.Rows) {
            if ($candidateRow.Cells['Cover'].Value -eq $true) {
                $hasCover = $true
                break
            }
        }
        if (-not $hasCover) {
            & $setCoverRow $eventArgs.RowIndex
        }
    })

    $getPlanFromGrid = {
        $chapters = @()
        $selectedCoverSource = ''
        foreach ($row in $grid.Rows) {
            if ($row.Cells['Cover'].Value -eq $true) {
                $selectedCoverSource = [string]$row.Cells['SourceFolder'].Value
            }
            $chapters += [pscustomobject][ordered]@{
                number = [string]$row.Cells['Number'].Value
                title = [string]$row.Cells['Title'].Value
                sourceFolder = [string]$row.Cells['SourceFolder'].Value
                sourceChapter = [string]$row.Cells['SourceChapter'].Value
                start = [string]$row.Cells['Start'].Value
                end = [string]$row.Cells['End'].Value
                mergeWithPrevious = ($row.Cells['Merge'].Value -eq $true)
                chapterCoverMode = if ([string]::IsNullOrWhiteSpace([string]$row.Cells['ChapterCoverMode'].Value)) { 'first' } else { [string]$row.Cells['ChapterCoverMode'].Value }
                chapterCoverPath = [string]$row.Cells['ChapterCoverPath'].Value
            }
        }
        if ([string]::IsNullOrWhiteSpace($selectedCoverSource) -and $grid.Rows.Count -gt 0) {
            $selectedCoverSource = [string]$grid.Rows[0].Cells['SourceFolder'].Value
        }
        return [pscustomobject][ordered]@{
            schemaVersion = 7
            outputName = $outputName.Text.Trim()
            coverSource = $selectedCoverSource
            customCoverPath = $script:OrganizerCustomCoverPath
            showChapterCovers = [bool]$showChapterCovers.Checked
            descriptionSource = $script:OrganizerDescriptionSource
            description = $script:OrganizerDescription
            descriptionLocked = $script:OrganizerDescriptionLocked
            selectedSourceFolders = @($sourceList.CheckedItems | ForEach-Object { [string]$_ })
            chapters = $chapters
        }
    }

    $setGridFromPlan = {
        param([object]$Plan)
        $planChapters = @((Get-ObjectProperty -Object $Plan -Name 'chapters' -Default @()))
        $sourceEntriesCache = @{}
        $chapterCoverUiState.Changing = $true
        $grid.SuspendLayout()
        $grid.Rows.Clear()
        $planOutputName = [string](Get-ObjectProperty -Object $Plan -Name 'outputName' -Default '')
        $outputName.Text = $planOutputName
        $sourceNames = @((Get-ObjectProperty -Object $Plan -Name 'selectedSourceFolders' -Default @()) | ForEach-Object {
            [string]$_
        } | Where-Object { $_ } | Select-Object -Unique)
        if ($sourceNames.Count -eq 0) {
            $sourceNames = @(@((Get-ObjectProperty -Object $Plan -Name 'chapters' -Default @())) | ForEach-Object {
                [string](Get-ObjectProperty -Object $_ -Name 'sourceFolder' -Default '')
            } | Where-Object { $_ } | Select-Object -Unique)
        }
        $requestedCover = [string](Get-ObjectProperty -Object $Plan -Name 'coverSource' -Default '')
        $script:OrganizerCustomCoverPath = [string](Get-ObjectProperty -Object $Plan -Name 'customCoverPath' -Default '')
        $showChapterCovers.Checked = [bool](Get-ObjectProperty -Object $Plan -Name 'showChapterCovers' -Default $false)
        & $updateWholeCoverButton
        $descriptionProperty = $Plan.PSObject.Properties['description']
        $script:OrganizerDescriptionSource = [string](Get-ObjectProperty -Object $Plan -Name 'descriptionSource' -Default '')
        if ($null -ne $descriptionProperty) {
            $script:OrganizerDescription = [string]$descriptionProperty.Value
            $script:OrganizerDescriptionLocked = [bool](Get-ObjectProperty -Object $Plan -Name 'descriptionLocked' -Default $true)
        }
        else {
            $script:OrganizerDescriptionSource = ''
            $script:OrganizerDescription = ''
            $script:OrganizerDescriptionLocked = $false
        }
        & $updateDescriptionButton
        try {
            for ($chapterIndex = 0; $chapterIndex -lt $planChapters.Count; $chapterIndex++) {
                $chapter = $planChapters[$chapterIndex]
                $status.Text = ('正在载入方案：第 {0}/{1} 行……' -f ($chapterIndex + 1), $planChapters.Count)
                $status.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
                $sourceFolder = [string](Get-ObjectProperty -Object $chapter -Name 'sourceFolder' -Default '')
                $sourceChapter = [string](Get-ObjectProperty -Object $chapter -Name 'sourceChapter' -Default '')
                $total = ''
                $sourceInfo = $null
                try {
                    $sourceInfo = Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $sourceFolder -SourceChapter $sourceChapter -SourceEntriesCache $sourceEntriesCache
                    $total = $sourceInfo.Count
                }
                catch {}
                $newRowIndex = $grid.Rows.Add(
                    [string](Get-ObjectProperty -Object $chapter -Name 'number' -Default ''),
                    [string](Get-ObjectProperty -Object $chapter -Name 'title' -Default ''),
                    $sourceFolder,
                    $sourceChapter,
                    [string](Get-ObjectProperty -Object $chapter -Name 'start' -Default ''),
                    [string](Get-ObjectProperty -Object $chapter -Name 'end' -Default ''),
                    [string]$total,
                    $(switch ([string](Get-ObjectProperty -Object $chapter -Name 'chapterCoverMode' -Default 'first')) { 'metadata' { '跟随元数据' } 'custom' { '自选' } 'none' { '隐藏' } default { '首图' } }),
                    [string](Get-ObjectProperty -Object $chapter -Name 'chapterCoverMode' -Default 'first'),
                    [string](Get-ObjectProperty -Object $chapter -Name 'chapterCoverPath' -Default ''),
                    [bool](Get-ObjectProperty -Object $chapter -Name 'mergeWithPrevious' -Default $false),
                    $false
                )
                $sourceComicPath = if ([string]::IsNullOrWhiteSpace($sourceFolder)) { '' } else { Join-Path $LibraryRoot $sourceFolder }
                & $setMetadataCoverCell $grid.Rows[$newRowIndex] $sourceInfo $sourceComicPath
                & $applyMergedRowDisplay $grid.Rows[$newRowIndex]
            }
        }
        finally {
            $grid.ResumeLayout()
            $chapterCoverUiState.Changing = $false
        }
        if ($grid.Rows.Count -gt 0) {
            $coverRowIndex = 0
            for ($index = 0; $index -lt $grid.Rows.Count; $index++) {
                if ([string]$grid.Rows[$index].Cells['SourceFolder'].Value -ceq $requestedCover) {
                    $coverRowIndex = $index
                    break
                }
            }
            & $setCoverRow $coverRowIndex
        }
        for ($index = 0; $index -lt $sourceList.Items.Count; $index++) {
            $sourceList.SetItemChecked($index, ($sourceNames -contains [string]$sourceList.Items[$index]))
        }
        & $refreshOutputNameChoices $sourceNames $planOutputName
        if (-not $script:OrganizerDescriptionLocked) { & $syncDefaultDescriptionFromOutputName }
        $script:lastDefaultOutput = $outputName.Text
        $script:OrganizerLoadedOnce = $true
        $loadSelected.Text = '重新载入所选文件夹'
        $loadNewSources.Enabled = $true
        $status.Text = ('已加载方案：{0} 话。' -f $grid.Rows.Count)
    }

    $getLoadedSourceNames = {
        return @($grid.Rows | ForEach-Object { [string]$_.Cells['SourceFolder'].Value } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    }

    $addSourceRows = {
        param([string[]]$Names, [bool]$ResetRows)
        $originalRowCount = $grid.Rows.Count
        if ($ResetRows) {
            $grid.Rows.Clear()
            $originalRowCount = 0
        }
        $existingSources = @(& $getLoadedSourceNames)
        $allSourceNames = @($existingSources + $Names | Select-Object -Unique)
        $outputNumber = $grid.Rows.Count
        $metadataOrderSources = @()
        $orderWarnings = @()
        try {
            foreach ($name in $Names) {
                $comicPath = Join-Path $LibraryRoot $name
                $sourceChapterEntries = @(Get-SourceChapterEntries -ComicPath $comicPath)
                if ($sourceChapterEntries.Count -gt 0 -and [string]$sourceChapterEntries[0].OrderSource -eq 'Metadata') {
                    $metadataOrderSources += $name
                }
                $sourceOrderWarning = @($sourceChapterEntries | ForEach-Object { [string]$_.OrderWarning } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
                foreach ($warning in $sourceOrderWarning) { $orderWarnings += ($name + '：' + $warning) }
                foreach ($chapterDirectory in $sourceChapterEntries) {
                    $imageCount = @($chapterDirectory.Images).Count
                    if ($imageCount -eq 0) { throw ($name + '\' + $chapterDirectory.Name + '：没有可载入的图片。') }
                    $outputNumber++
                    $initialFields = Get-InitialOrganizerChapterFields -ChapterName $chapterDirectory.Name -IsRootChapter ([bool]$chapterDirectory.IsRootChapter) -DefaultNumber ([string]$outputNumber) -PreserveNumericNumber ($allSourceNames.Count -eq 1)
                    $hasConfiguredCover = $null -ne $chapterDirectory.PSObject.Properties['ConfiguredCoverMode']
                    $chapterCoverMode = if ($hasConfiguredCover) { 'metadata' } else { 'first' }
                    $chapterCoverPath = ''
                    if ($null -ne $chapterDirectory.PSObject.Properties['ConfiguredShowChapterCovers'] -and [bool]$chapterDirectory.ConfiguredShowChapterCovers) {
                        $showChapterCovers.Checked = $true
                    }
                    $chapterCoverDisplay = switch ($chapterCoverMode) {
                        'metadata' { '跟随元数据' }
                        'custom' { '自选' }
                        'none' { '隐藏' }
                        default { '首图' }
                    }
                    $newRowIndex = $grid.Rows.Add($initialFields.Number, $initialFields.Title, $name, $chapterDirectory.Name, 1, $imageCount, $imageCount, $chapterCoverDisplay, $chapterCoverMode, $chapterCoverPath, $false, ($grid.Rows.Count -eq 0))
                    & $setMetadataCoverCell $grid.Rows[$newRowIndex] $chapterDirectory $comicPath
                    & $applyMergedRowDisplay $grid.Rows[$newRowIndex]
                }
            }
        }
        catch {
            while ($grid.Rows.Count -gt $originalRowCount) { $grid.Rows.RemoveAt($grid.Rows.Count - 1) }
            & $showMessage $_.Exception.Message '来源载入失败' ([System.Windows.Forms.MessageBoxIcon]::Error)
            return $null
        }
        return [pscustomobject]@{
            AddedSources = @($Names)
            AddedRows = $grid.Rows.Count - $originalRowCount
            MetadataOrderSources = @($metadataOrderSources)
            OrderWarnings = @($orderWarnings)
        }
    }

    $showOrderWarnings = {
        param([object]$LoadResult)
        if ($null -ne $LoadResult -and @($LoadResult.OrderWarnings).Count -gt 0) {
            & $showMessage (@($LoadResult.OrderWarnings) -join "`r`n`r`n") '章节顺序配置未完全采用' ([System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }

    $loadSelected.add_Click({
        $selectedNames = @($sourceList.CheckedItems | ForEach-Object { [string]$_ })
        if ($selectedNames.Count -eq 0) {
            & $showMessage '请至少勾选一个来源漫画文件夹。' '尚未选择' ([System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        if ($script:OrganizerLoadedOnce -and $grid.Rows.Count -gt 0) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                $form,
                ('“重新载入所选文件夹”会重建右侧表格并重置当前编辑。' + "`r`n`r`n" + '如果只是刚勾选了新文件夹，请选择【否】，改用“载入新增文件夹”，原有编辑就会完整保留。' + "`r`n`r`n" + '仍要重新载入吗？'),
                '二次确认：重新载入会重置编辑',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        & $pushUndoSnapshot
        $currentOutputText = $outputName.Text.Trim()
        $useDefaultName = [string]::IsNullOrWhiteSpace($currentOutputText) -or $currentOutputText -ceq $script:lastDefaultOutput
        $wasLoaded = $script:OrganizerLoadedOnce
        $loadResult = & $addSourceRows $selectedNames $true
        if ($null -eq $loadResult) { return }
        $preferredName = if ($useDefaultName) { $selectedNames[0] } else { $currentOutputText }
        & $refreshOutputNameChoices $selectedNames $preferredName
        if (-not $script:OrganizerDescriptionLocked) { & $syncDefaultDescriptionFromOutputName }
        if ($useDefaultName) { $script:lastDefaultOutput = $selectedNames[0] }
        $script:OrganizerLoadedOnce = $true
        $loadSelected.Text = '重新载入所选文件夹'
        $loadNewSources.Enabled = $true
        $metadataSourceCount = @($loadResult.MetadataOrderSources).Count
        $orderStatus = if ($metadataSourceCount -gt 0) { '其中 {0} 个来源已采用元数据章节顺序。' -f $metadataSourceCount } else { '未发现可采用的元数据章节顺序。' }
        $status.Text = if ($wasLoaded) {
            '已重新载入 {0} 个来源、{1} 行章节；{2} 原表格编辑已重置。' -f $selectedNames.Count, $grid.Rows.Count, $orderStatus
        }
        else { '已载入 {0} 个来源、{1} 行章节；{2}' -f $selectedNames.Count, $grid.Rows.Count, $orderStatus }
        & $showOrderWarnings $loadResult
    })

    $loadNewSources.add_Click({
        $selectedNames = @($sourceList.CheckedItems | ForEach-Object { [string]$_ })
        $loadedNames = @(& $getLoadedSourceNames)
        $newNames = @($selectedNames | Where-Object { $loadedNames -notcontains $_ })
        if ($newNames.Count -eq 0) {
            & $showMessage '当前勾选项中没有尚未载入的新文件夹。原有表格编辑保持不变。' '没有新增来源' ([System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        & $pushUndoSnapshot
        $currentOutputText = $outputName.Text
        $loadResult = & $addSourceRows $newNames $false
        if ($null -eq $loadResult) { return }
        $allLoadedNames = @(& $getLoadedSourceNames)
        & $refreshOutputNameChoices $allLoadedNames $currentOutputText
        $status.Text = ('已追加 {0} 个新来源、{1} 行章节；原有 {2} 行编辑未重置。' -f $newNames.Count, $loadResult.AddedRows, ($grid.Rows.Count - $loadResult.AddedRows))
        & $showOrderWarnings $loadResult
    })

    $rescanSources.add_Click({
        & $pushUndoSnapshot
        $checkedNames = @($sourceList.CheckedItems | ForEach-Object { [string]$_ })
        $loadedNames = @(& $getLoadedSourceNames)
        $keepChecked = @($checkedNames + $loadedNames | Select-Object -Unique)
        try {
            $rescannedCandidates = @(Get-CandidateFolders -LibraryRoot $LibraryRoot)
            $sourceList.BeginUpdate()
            try {
                $sourceList.Items.Clear()
                foreach ($candidate in $rescannedCandidates) {
                    $index = $sourceList.Items.Add($candidate.Name)
                    if ($keepChecked -contains $candidate.Name) { $sourceList.SetItemChecked($index, $true) }
                }
            }
            finally { $sourceList.EndUpdate() }

            $refreshProblems = @()
            $refreshedRows = 0
            $sourceEntriesCache = @{}
            for ($rowIndex = 0; $rowIndex -lt $grid.Rows.Count; $rowIndex++) {
                $row = $grid.Rows[$rowIndex]
                $status.Text = ('正在刷新已载入章节：第 {0}/{1} 行……' -f ($rowIndex + 1), $grid.Rows.Count)
                $status.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
                $sourceFolder = [string]$row.Cells['SourceFolder'].Value
                $sourceChapter = [string]$row.Cells['SourceChapter'].Value
                $oldTotal = 0
                [void][int]::TryParse([string]$row.Cells['Total'].Value, [ref]$oldTotal)
                try {
                    $refreshedSourceInfo = Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $sourceFolder -SourceChapter $sourceChapter -SourceEntriesCache $sourceEntriesCache
                    $newTotal = $refreshedSourceInfo.Count
                    $endValue = 0
                    $endWasTotal = [int]::TryParse([string]$row.Cells['End'].Value, [ref]$endValue) -and $oldTotal -gt 0 -and $endValue -eq $oldTotal
                    if ($endWasTotal -or [string]::IsNullOrWhiteSpace([string]$row.Cells['End'].Value)) { $row.Cells['End'].Value = $newTotal }
                    $row.Cells['Total'].Value = $newTotal
                    & $setMetadataCoverCell $row $refreshedSourceInfo (Join-Path $LibraryRoot $sourceFolder)
                    $refreshedRows++
                }
                catch {
                    $row.Cells['Total'].Value = '缺失'
                    $refreshProblems += ($sourceFolder + '\' + $sourceChapter + '：' + $_.Exception.Message)
                }
            }
            & $refreshOutputNameChoices (@(& $getLoadedSourceNames)) $outputName.Text
            $loadNewSources.Enabled = $grid.Rows.Count -gt 0
            $status.Text = ('重新扫描完成：来源库 {0} 个文件夹，已刷新 {1} 行载入信息；表格编辑和手动范围保持不变。' -f $rescannedCandidates.Count, $refreshedRows)
            if ($refreshProblems.Count -gt 0) {
                & $showMessage ($refreshProblems -join "`r`n") '部分已载入来源发生变化' ([System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
        catch {
            & $showMessage $_.Exception.Message '重新扫描失败' ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $getRowValues = {
        param([System.Windows.Forms.DataGridViewRow]$Row)
        $values = @()
        for ($columnIndex = 0; $columnIndex -lt $grid.Columns.Count; $columnIndex++) { $values += ,$Row.Cells[$columnIndex].Value }
        return $values
    }
    $setRowValues = {
        param([System.Windows.Forms.DataGridViewRow]$Row, [object[]]$Values)
        for ($columnIndex = 0; $columnIndex -lt $grid.Columns.Count; $columnIndex++) { $Row.Cells[$columnIndex].Value = $Values[$columnIndex] }
        & $applyMergedRowDisplay $Row
    }

    $duplicateRowsAction = {
        $rows = @(& $getBatchRows | Sort-Object Index)
        if ($rows.Count -eq 0) { return }
        & $pushUndoSnapshot
        $insertIndex = $rows[-1].Index + 1
        $firstInserted = $insertIndex
        foreach ($sourceRow in $rows) {
            $values = @(& $getRowValues $sourceRow)
            $values[$grid.Columns['Number'].Index] = ''
            $values[$grid.Columns['Merge'].Index] = $false
            $values[$grid.Columns['Cover'].Index] = $false
            $values[$grid.Columns['Batch'].Index] = $false
            $grid.Rows.Insert($insertIndex, $values)
            & $applyMergedRowDisplay $grid.Rows[$insertIndex]
            $insertIndex++
        }
        $grid.CurrentCell = $grid.Rows[$firstInserted].Cells['Number']
        $status.Text = ('已按原相对顺序复制 {0} 行；副本话序已留空，且不会自动并入或成为整本封面。' -f $rows.Count)
    }
    $duplicateRow.add_Click({ & $duplicateRowsAction })

    $deleteRowsAction = {
        $rows = @(& $getBatchRows | Sort-Object Index)
        if ($rows.Count -eq 0) { return }
        & $pushUndoSnapshot
        $firstDeletedIndex = $rows[0].Index
        $deletedCover = @($rows | Where-Object { $_.Cells['Cover'].Value -eq $true }).Count -gt 0
        foreach ($row in @($rows | Sort-Object Index -Descending)) { $grid.Rows.RemoveAt($row.Index) }
        if ($deletedCover -and $grid.Rows.Count -gt 0) { & $setCoverRow ([Math]::Min($firstDeletedIndex, $grid.Rows.Count - 1)) }
        $status.Text = ('已删除 {0} 行；其余行话序保持不变。' -f $rows.Count)
    }
    $deleteRow.add_Click({ & $deleteRowsAction })

    $mergeSelectedAction = {
        $rows = @(& $getBatchRows | Sort-Object Index)
        if ($rows.Count -lt 2) {
            & $showMessage '请在第一列至少勾选两行。所选行可以不连续。' '合并为同一话' ([System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        & $pushUndoSnapshot
        $baseIndex = $rows[0].Index
        $coverOffset = -1
        $storedRows = @()
        for ($offset = 0; $offset -lt $rows.Count; $offset++) {
            if ($rows[$offset].Cells['Cover'].Value -eq $true) { $coverOffset = $offset }
            $storedRows += ,@(& $getRowValues $rows[$offset])
        }
        $chapterCoverUiState.Changing = $true
        $coverSelectionState.Changing = $true
        try {
            foreach ($row in @($rows | Sort-Object Index -Descending)) { $grid.Rows.RemoveAt($row.Index) }
            for ($offset = 0; $offset -lt $storedRows.Count; $offset++) {
                $values = @($storedRows[$offset])
                $values[$grid.Columns['Merge'].Index] = ($offset -gt 0)
                $values[$grid.Columns['Batch'].Index] = ($offset -eq 0)
                $values[$grid.Columns['Cover'].Index] = ($offset -eq $coverOffset)
                $grid.Rows.Insert($baseIndex + $offset, $values)
                & $applyMergedRowDisplay $grid.Rows[$baseIndex + $offset]
            }
        }
        finally {
            $coverSelectionState.Changing = $false
            $chapterCoverUiState.Changing = $false
        }
        if ($coverOffset -ge 0) { & $setCoverRow ($baseIndex + $coverOffset) }
        $grid.CurrentCell = $grid.Rows[$baseIndex].Cells['Number']
        $status.Text = ('已把 {0} 个所选来源行合并为同一话；并入行的原话序、章节名和封面信息会保留并灰显，但逻辑上不单独生效。现在可直接点击“从选中行后续编号”。' -f $rows.Count)
    }
    $mergeSelected.add_Click({ & $mergeSelectedAction })

    $splitRow.add_Click({
        if ($null -eq $grid.CurrentRow) {
            & $showMessage '请先选中需要拆分的总集行。' '尚未选择' ([System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $row = $grid.CurrentRow
        $defaultRange = ([string]$row.Cells['Start'].Value) + '-' + ([string]$row.Cells['End'].Value)
        $input = [Microsoft.VisualBasic.Interaction]::InputBox(
            "输入每话的图片范围，起止均包含。`r`n示例：1-30,31-62,63-90",
            '按范围拆分总集',
            $defaultRange
        )
        if ([string]::IsNullOrWhiteSpace($input)) { return }
        $ranges = @()
        $sourceTotal = 0
        [void][int]::TryParse([string]$row.Cells['Total'].Value, [ref]$sourceTotal)
        foreach ($part in @($input -split '[,，;；]+')) {
            if ($part -notmatch '^\s*(\d+)\s*(?:-|~|—|至)\s*(\d+)\s*$') {
                & $showMessage ('无法识别范围：' + $part) '范围格式错误' ([System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            $startValue = [int]$matches[1]
            $endValue = [int]$matches[2]
            if ($startValue -lt 1 -or $endValue -lt $startValue) {
                & $showMessage ('无效范围：' + $part) '范围格式错误' ([System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            if ($sourceTotal -gt 0 -and $endValue -gt $sourceTotal) {
                & $showMessage ('范围超出来源总数 ' + $sourceTotal + '：' + $part) '范围格式错误' ([System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            $ranges += ,@($startValue, $endValue)
        }
        $simpleBoundaryIssues = @()
        for ($rangeIndex = 1; $rangeIndex -lt $ranges.Count; $rangeIndex++) {
            $previousRange = $ranges[$rangeIndex - 1]
            $currentRange = $ranges[$rangeIndex]
            if ([int]$currentRange[0] -eq [int]$previousRange[1] -and ([int]$currentRange[0] + 1) -le [int]$currentRange[1]) {
                $simpleBoundaryIssues += ('第 {0} 段 {1}-{2} 与前一段重复了图片 {1}，建议改为 {3}-{2}' -f ($rangeIndex + 1), $currentRange[0], $currentRange[1], ([int]$currentRange[0] + 1))
            }
        }
        $simpleOverlapKept = $false
        if ($simpleBoundaryIssues.Count -gt 0) {
            $simpleAnswer = [System.Windows.Forms.MessageBox]::Show(
                $form,
                ("发现很可能是手误的相邻边界重复：`r`n`r`n{0}`r`n`r`n选择【是】：自动把后一段起点加 1 后继续。`r`n选择【否】：保留重复并继续。`r`n选择【取消】：返回修改。" -f ($simpleBoundaryIssues -join "`r`n")),
                '发现可自动修正的范围错误',
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($simpleAnswer -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
            if ($simpleAnswer -eq [System.Windows.Forms.DialogResult]::Yes) {
                for ($rangeIndex = 1; $rangeIndex -lt $ranges.Count; $rangeIndex++) {
                    if ([int]$ranges[$rangeIndex][0] -eq [int]$ranges[$rangeIndex - 1][1] -and ([int]$ranges[$rangeIndex][0] + 1) -le [int]$ranges[$rangeIndex][1]) {
                        $ranges[$rangeIndex][0] = [int]$ranges[$rangeIndex][0] + 1
                    }
                }
            }
            else { $simpleOverlapKept = $true }
        }

        $originalStart = 0
        $originalEnd = 0
        [void][int]::TryParse([string]$row.Cells['Start'].Value, [ref]$originalStart)
        [void][int]::TryParse([string]$row.Cells['End'].Value, [ref]$originalEnd)
        $continuityIssues = @()
        if ($ranges.Count -gt 0 -and $originalStart -gt 0 -and [int]$ranges[0][0] -ne $originalStart) {
            $continuityIssues += ('第一段从 {0} 开始，但原行从 {1} 开始。' -f $ranges[0][0], $originalStart)
        }
        for ($rangeIndex = 1; $rangeIndex -lt $ranges.Count; $rangeIndex++) {
            $expectedStart = [int]$ranges[$rangeIndex - 1][1] + 1
            $actualStart = [int]$ranges[$rangeIndex][0]
            if ($actualStart -gt $expectedStart) {
                $continuityIssues += ('第 {0} 段之前有缺口：缺少 {1}-{2}。' -f ($rangeIndex + 1), $expectedStart, ($actualStart - 1))
            }
            elseif ($actualStart -lt $expectedStart) {
                $continuityIssues += ('第 {0} 段与前面重叠：重复 {1}-{2}。' -f ($rangeIndex + 1), $actualStart, ([Math]::Min([int]$ranges[$rangeIndex][1], [int]$ranges[$rangeIndex - 1][1])))
            }
        }
        if ($ranges.Count -gt 0 -and $originalEnd -gt 0 -and [int]$ranges[-1][1] -ne $originalEnd) {
            $continuityIssues += ('最后一段结束于 {0}，但原行结束于 {1}。' -f $ranges[-1][1], $originalEnd)
        }
        if ($continuityIssues.Count -gt 0 -and -not ($simpleOverlapKept -and $continuityIssues.Count -eq $simpleBoundaryIssues.Count)) {
            $continueAnswer = [System.Windows.Forms.MessageBox]::Show(
                $form,
                ("这些范围没有把原行每一张图片连续且仅使用一次：`r`n`r`n{0}`r`n`r`n这可能是有意舍弃或重复。仍按当前范围拆分吗？" -f ($continuityIssues -join "`r`n")),
                '二次确认：范围并非连续完整分割',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($continueAnswer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        & $pushUndoSnapshot
        $insertIndex = $row.Index
        $sourceFolderValue = [string]$row.Cells['SourceFolder'].Value
        $sourceChapterValue = [string]$row.Cells['SourceChapter'].Value
        $totalValue = [string]$row.Cells['Total'].Value
        $originalNumber = [string]$row.Cells['Number'].Value
        $wasMerged = ($row.Cells['Merge'].Value -eq $true)
        $wasCover = ($row.Cells['Cover'].Value -eq $true)
        $originalChapterCover = [string]$row.Cells['ChapterCover'].Value
        $originalChapterCoverMode = [string]$row.Cells['ChapterCoverMode'].Value
        $originalChapterCoverPath = [string]$row.Cells['ChapterCoverPath'].Value
        $originalMetadataCoverInfo = [string]$row.Cells['MetadataCoverInfo'].Value
        $originalMetadataCoverTooltip = [string]$row.Cells['MetadataCoverInfo'].ToolTipText
        $grid.Rows.RemoveAt($insertIndex)
        for ($offset = 0; $offset -lt $ranges.Count; $offset++) {
            $range = $ranges[$offset]
            $grid.Rows.Insert(
                $insertIndex + $offset,
                @(
                    $(if ($offset -eq 0) { $originalNumber } else { '' }), '', $sourceFolderValue, $sourceChapterValue,
                    $range[0], $range[1], $totalValue,
                    $(if ($offset -eq 0) { $originalChapterCover } else { '首图' }),
                    $(if ($offset -eq 0) { $originalChapterCoverMode } else { 'first' }),
                    $(if ($offset -eq 0) { $originalChapterCoverPath } else { '' }),
                    ($wasMerged -and $offset -eq 0), ($wasCover -and $offset -eq 0), $true, $originalMetadataCoverInfo
                )
            )
            $grid.Rows[$insertIndex + $offset].Cells['MetadataCoverInfo'].ToolTipText = $originalMetadataCoverTooltip
            $grid.Rows[$insertIndex + $offset].Cells['ChapterCover'].ToolTipText = '跟随元数据时：' + $originalMetadataCoverTooltip
            & $applyMergedRowDisplay $grid.Rows[$insertIndex + $offset]
        }
        if ($wasCover) { & $setCoverRow $insertIndex }
        $grid.CurrentCell = $grid.Rows[$insertIndex].Cells['Number']
        $status.Text = ('已将一行拆成 {0} 话；请填写新增行话序/特殊标签（如 4.5、特典话），或选中数字锚点后点击“从选中行后续编号”。' -f $ranges.Count)
    })

    $moveRow = {
        param([int]$Delta)
        $hadBatchSelection = @($grid.Rows | Where-Object { $_.Cells['Batch'].Value -eq $true }).Count -gt 0
        $rows = @(& $getBatchRows)
        if ($rows.Count -eq 0) { return }
        & $pushUndoSnapshot
        if (-not $hadBatchSelection) { $rows[0].Cells['Batch'].Value = $true }
        $chapterCoverUiState.Changing = $true
        $coverSelectionState.Changing = $true
        try {
            if ($Delta -lt 0) {
                for ($index = 1; $index -lt $grid.Rows.Count; $index++) {
                    if ($grid.Rows[$index].Cells['Batch'].Value -eq $true -and $grid.Rows[$index - 1].Cells['Batch'].Value -ne $true) {
                        $currentValues = @(& $getRowValues $grid.Rows[$index])
                        $previousValues = @(& $getRowValues $grid.Rows[$index - 1])
                        & $setRowValues $grid.Rows[$index - 1] $currentValues
                        & $setRowValues $grid.Rows[$index] $previousValues
                    }
                }
            }
            else {
                for ($index = $grid.Rows.Count - 2; $index -ge 0; $index--) {
                    if ($grid.Rows[$index].Cells['Batch'].Value -eq $true -and $grid.Rows[$index + 1].Cells['Batch'].Value -ne $true) {
                        $currentValues = @(& $getRowValues $grid.Rows[$index])
                        $nextValues = @(& $getRowValues $grid.Rows[$index + 1])
                        & $setRowValues $grid.Rows[$index + 1] $currentValues
                        & $setRowValues $grid.Rows[$index] $nextValues
                    }
                }
            }
        }
        finally {
            $coverSelectionState.Changing = $false
            $chapterCoverUiState.Changing = $false
        }
        $newRows = @($grid.Rows | Where-Object { $_.Cells['Batch'].Value -eq $true } | Sort-Object Index)
        if ($newRows.Count -gt 0) { $grid.CurrentCell = $newRows[0].Cells['Number'] }
        $status.Text = ('已将 {0} 个所选行整体{1}一格；相对顺序与话序标签保持不变。' -f $newRows.Count, $(if ($Delta -lt 0) { '上移' } else { '下移' }))
        if (-not $hadBatchSelection) { foreach ($row in $newRows) { $row.Cells['Batch'].Value = $false } }
    }
    $moveRowsToEdge = {
        param([ValidateSet('Top', 'Bottom')][string]$Destination)
        $hadBatchSelection = @($grid.Rows | Where-Object { $_.Cells['Batch'].Value -eq $true }).Count -gt 0
        $rows = @(& $getBatchRows | Sort-Object Index)
        if ($rows.Count -eq 0) { return }
        & $pushUndoSnapshot
        if (-not $hadBatchSelection) { $rows[0].Cells['Batch'].Value = $true }
        $storedRows = @($rows | ForEach-Object { ,@(& $getRowValues $_) })
        $chapterCoverUiState.Changing = $true
        $coverSelectionState.Changing = $true
        try {
            foreach ($row in @($rows | Sort-Object Index -Descending)) { $grid.Rows.RemoveAt($row.Index) }
            $insertIndex = if ($Destination -eq 'Top') { 0 } else { $grid.Rows.Count }
            foreach ($values in $storedRows) {
                $grid.Rows.Insert($insertIndex, @($values))
                & $applyMergedRowDisplay $grid.Rows[$insertIndex]
                $insertIndex++
            }
        }
        finally {
            $coverSelectionState.Changing = $false
            $chapterCoverUiState.Changing = $false
        }
        $newRows = @($grid.Rows | Where-Object { $_.Cells['Batch'].Value -eq $true } | Sort-Object Index)
        if ($newRows.Count -gt 0) { $grid.CurrentCell = $newRows[0].Cells['Number'] }
        $status.Text = ('已将 {0} 个所选行整体移到最{1}；相对顺序与话序标签保持不变。' -f $newRows.Count, $(if ($Destination -eq 'Top') { '上' } else { '下' }))
        if (-not $hadBatchSelection) { foreach ($row in $newRows) { $row.Cells['Batch'].Value = $false } }
    }
    $moveTop.add_Click({ & $moveRowsToEdge 'Top' })
    $moveUp.add_Click({ & $moveRow -1 })
    $moveDown.add_Click({ & $moveRow 1 })
    $moveBottom.add_Click({ & $moveRowsToEdge 'Bottom' })
    $autoNumber.add_Click({ & $renumberRows })
    $selectAllRows.add_Click({ foreach ($row in $grid.Rows) { $row.Cells['Batch'].Value = $true }; $status.Text = ('已选择全部 {0} 行。' -f $grid.Rows.Count) })
    $clearSelectedRows.add_Click({ foreach ($row in $grid.Rows) { $row.Cells['Batch'].Value = $false }; $status.Text = '已清空第一列行选择。' })
    $clearWorkspace.add_Click({
        if ($grid.Rows.Count -eq 0 -and [string]::IsNullOrWhiteSpace($outputName.Text)) {
            $status.Text = '右侧整理内容已经是空的。'
            return
        }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            $form,
            "确定清空右侧全部整理内容吗？`r`n`r`n这会清除当前表格、输出名称、简介和封面设置，但不会删除任何原漫画文件，也不会清空左侧来源文件夹库。",
            '二次确认：清空右侧内容',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        & $pushUndoSnapshot
        $chapterCoverUiState.Changing = $true
        $coverSelectionState.Changing = $true
        try {
            $grid.Rows.Clear()
            $outputName.Items.Clear()
            $outputName.Text = ''
            for ($itemIndex = 0; $itemIndex -lt $sourceList.Items.Count; $itemIndex++) { $sourceList.SetItemChecked($itemIndex, $false) }
            $showChapterCovers.Checked = $false
            $script:lastDefaultOutput = ''
            $script:OrganizerDescription = ''
            $script:OrganizerDescriptionSource = ''
            $script:OrganizerDescriptionLocked = $false
            $script:OrganizerCustomCoverPath = ''
            $script:OrganizerLoadedOnce = $false
            $loadSelected.Text = '载入所选文件夹'
            $loadNewSources.Enabled = $false
        }
        finally {
            $coverSelectionState.Changing = $false
            $chapterCoverUiState.Changing = $false
        }
        & $updateDescriptionButton
        & $updateWholeCoverButton
        $status.Text = '右侧整理内容已清空，可以开始整理下一部漫画。左侧来源文件夹库仍保留。'
    })

    $helpButton.add_Click({
        $helpText = @'
1. 首次勾选来源后点击“载入所选文件夹”；载入后该按钮会变为“重新载入”，使用它会先二次确认，再重建表格并重置现有编辑。
2. 编辑中途新增来源时，先“重新扫描来源文件夹库”，再勾选新来源并点击“载入新增文件夹”；它只追加新来源，原有行、范围、命名与排序都不会重置。
3. 输出漫画名称可手动输入，也可用右侧下拉箭头直接选择已载入漫画名称。
4. 简介默认跟随输出漫画名称所选来源；在简介窗口按“确定”后即固定，不再随名称来源变化。
5. 来源若有元数据.json 且 chapterInfos 完整有效，会按其中的 order 排列；无配置或无法完整匹配时按名称自然排序并提示原因。
6. 根目录若为 P01_001、P02_001 等格式，会按前缀自动生成多行章节；无法识别页码时按名称自然排序。
7. 表格从上到下就是阅读顺序，可修改话序、章节名、范围并上下移动；非数字话序会保留原特殊名称。
8. 第一列“选择”用于批量操作；封面、复制、删除、上移、下移、移到最上和移到最下等按钮优先处理勾选行，未勾选时处理当前行。
9. “合并所选为同一话”允许所选行不连续；整理器会把它们聚拢到第一条所选行的位置并按原相对顺序合并。并入行会保留原话序、章节名、章节封面和元数据指向并灰显，便于辨认来源，但这些字段逻辑上不会单独生效或占用章节编号。
10. “按范围拆分选中行”会检查是否连续完整覆盖；1-3,3-10 这类边界重复可自动修正，其他缺口或重叠必须二次确认。
11. “整本封面”可沿用勾选来源，也可选择任意本地图片；自选图片只会复制，不会改动原文件。
12. “漫画目录显示每话封面”是整本总开关：未勾选时目录不显示任何章节缩略图；勾选后，每行的跟随元数据、首图、自选或隐藏设置才会生效。“元数据实际指向”列会显示跟随后最终是首图、隐藏还是资源目录中的自选图片。
13. 图片被重复使用不再直接报错中止；整理前会汇总重叠范围并二次确认，疑似简单边界手误会单独标明。
14. 原文件不会修改；结果输出到“整理完成”，正文统一重命名为 0001、0002……。
15. “从选中行后续编号”只计算合并后的独立逻辑章节，并入同一话的来源行不会占号，也不会改写其灰显的原话序；连续勾选多个逻辑章节时以最后一个为锚点。
16. “清空右侧内容”会二次确认，只清空当前整理方案，不删除原漫画，也保留左侧来源文件夹库。
17. 纯数字或复合页码会检查重复和缺号；无法识别的名称排序模式会明确警告无法判断缺图。
18. “撤销 / 恢复”会保存最近 40 次表格与界面操作，包括行内容、顺序、简介、整本封面、章节封面和载入状态；也可使用 Ctrl+Z / Ctrl+Y。
'@
        & $showMessage $helpText '漫画整理器使用说明' ([System.Windows.Forms.MessageBoxIcon]::Information)
    })

    $validateButton.add_Click({
        try {
            $plan = & $getPlanFromGrid
            $audit = Test-OrganizerPlan -Plan $plan -LibraryRoot $LibraryRoot -IgnoreExistingOutput
            if (-not $audit.IsValid) {
                & $showMessage (($audit.Errors | Select-Object -First 30) -join "`r`n") '方案未通过' ([System.Windows.Forms.MessageBoxIcon]::Error)
                $status.Text = ('方案有 {0} 个问题。' -f $audit.Errors.Count)
                return
            }
            $warningText = ''
            if ($audit.Warnings.Count -gt 0) { $warningText = "`r`n`r`n警告：`r`n" + ($audit.Warnings -join "`r`n") }
            & $showMessage ("核验通过。`r`n{0} 话，{1} 张图片。`r`n输出：{2}{3}" -f $audit.ChapterCount, $audit.TotalImages, $audit.OutputPath, $warningText) '方案有效' ([System.Windows.Forms.MessageBoxIcon]::Information)
            $status.Text = ('核验通过：{0} 话，{1} 张图片。' -f $audit.ChapterCount, $audit.TotalImages)
        }
        catch {
            & $showMessage $_.Exception.Message '核验失败' ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $savePlan.add_Click({
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Title = '保存漫画整理方案'
        $dialog.InitialDirectory = $LibraryRoot
        $dialog.Filter = 'JSON 方案 (*.json)|*.json'
        $safeName = $outputName.Text -replace '[<>:"/\\|?*]', '_'
        if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = '漫画整理方案' }
        $dialog.FileName = $safeName + ' - 整理方案.json'
        if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                Write-JsonFile -Path $dialog.FileName -Value (& $getPlanFromGrid)
                $status.Text = ('方案已保存：' + $dialog.FileName)
            }
            catch {
                & $showMessage $_.Exception.Message '保存失败' ([System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    $loadPlan.add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = '加载漫画整理方案'
        $dialog.InitialDirectory = $LibraryRoot
        $dialog.Filter = 'JSON 方案 (*.json)|*.json'
        if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($script:OrganizerIsRunning) { return }
            & $pushUndoSnapshot
            & $setOrganizerBusy $true
            try {
                & $setGridFromPlan (Read-OrganizerPlan -Path $dialog.FileName)
            }
            catch {
                & $showMessage $_.Exception.Message '加载失败' ([System.Windows.Forms.MessageBoxIcon]::Error)
            }
            finally { & $setOrganizerBusy $false }
        }
    })

    $generateButton.add_Click({
        if ($script:OrganizerIsRunning) { return }
        & $setOrganizerBusy $true
        try {
            $status.Text = '正在读取界面中的整理方案…'
            [System.Windows.Forms.Application]::DoEvents()
            $plan = & $getPlanFromGrid
            $audit = Test-OrganizerPlan -Plan $plan -LibraryRoot $LibraryRoot -ProgressPrefix '正在核验整理方案'
            if (-not $audit.IsValid) {
                $status.Text = '方案未通过核验，请查看提示后修改。'
                & $showMessage (($audit.Errors | Select-Object -First 30) -join "`r`n") '方案未通过' ([System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            if ($audit.OverlapWarnings.Count -gt 0) {
                $overlapLines = @($audit.OverlapWarnings | Select-Object -First 20)
                $overlapText = $overlapLines -join "`r`n"
                if ($audit.OverlapWarnings.Count -gt $overlapLines.Count) {
                    $overlapText += ("`r`n……另有 {0} 项未显示。" -f ($audit.OverlapWarnings.Count - $overlapLines.Count))
                }
                $simpleText = if ($audit.SimpleOverlapWarnings.Count -gt 0) {
                    "`r`n`r`n其中检测到的简单边界错误：`r`n" + (@($audit.SimpleOverlapWarnings | Select-Object -First 10) -join "`r`n")
                }
                else { '' }
                $overlapAnswer = [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    ("以下图片被重复使用：`r`n`r`n{0}{1}`r`n`r`n重复有可能是你的有意安排。确认仍按当前方案继续吗？" -f $overlapText, $simpleText),
                    '二次确认：存在重复图片',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($overlapAnswer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }
            if ($audit.OmissionWarnings.Count -gt 0) {
                $omissionLines = @($audit.OmissionWarnings | Select-Object -First 20)
                $omissionText = $omissionLines -join "`r`n"
                if ($audit.OmissionWarnings.Count -gt $omissionLines.Count) {
                    $omissionText += ("`r`n……另有 {0} 项未显示。" -f ($audit.OmissionWarnings.Count - $omissionLines.Count))
                }
                $omissionAnswer = [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    ("以下来源内容没有被选择，因此不会出现在整理结果中：`r`n`r`n{0}`r`n`r`n这是你有意舍弃的内容吗？选择 是 后才会继续。" -f $omissionText),
                    '二次确认：舍弃未选图片',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($omissionAnswer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }
            $answer = [System.Windows.Forms.MessageBox]::Show(
                $form,
                ("即将复制整理：`r`n{0} 话，{1} 张图片。`r`n`r`n输出到：`r`n{2}`r`n`r`n原文件不会改动。是否继续？" -f $audit.ChapterCount, $audit.TotalImages, $audit.OutputPath),
                '确认开始整理',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            $resultPath = Invoke-OrganizerPlan -Audit $audit -LibraryRoot $LibraryRoot -ProgressPrefix '正在复制整理'
            $status.Text = ('整理完成：' + $resultPath)
            & $showMessage ("整理并复核完成。`r`n`r`n$resultPath`r`n`r`n确认无误后，可手动复制到漫画大文件夹。") '整理完成' ([System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            $status.Text = '整理失败：' + $_.Exception.Message
            & $showMessage $_.Exception.Message '整理失败' ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
        finally { & $setOrganizerBusy $false }
    })

    $form.add_FormClosing({
        param($sender, $eventArgs)
        if ($script:OrganizerIsRunning) {
            $eventArgs.Cancel = $true
            $status.Text = '整理任务仍在进行，请等待完成后再关闭窗口。'
        }
    })

    if ($SmokeTest) {
        $smokeCandidate = $null
        $smokeChapterCount = -1
        foreach ($candidate in $Candidates) {
            try {
                $candidateChapterCount = @(Get-SourceChapterEntries -ComicPath (Join-Path $LibraryRoot $candidate.Name)).Count
                if ($candidateChapterCount -gt $smokeChapterCount) {
                    $smokeCandidate = $candidate
                    $smokeChapterCount = $candidateChapterCount
                }
            }
            catch {}
        }
        if ($null -eq $smokeCandidate) { throw '整理器 UI 冒烟测试找不到可载入来源。' }
        $smokeSourceIndex = $sourceList.Items.IndexOf([string]$smokeCandidate.Name)
        if ($smokeSourceIndex -ge 0) { $sourceList.SetItemChecked($smokeSourceIndex, $true) }
        $smokeLoadResult = & $addSourceRows @([string]$smokeCandidate.Name) $true
        if ($null -eq $smokeLoadResult -or $grid.Rows.Count -eq 0) { throw '整理器 UI 冒烟测试无法载入章节行。' }
        if ([string]::IsNullOrWhiteSpace([string]$grid.Rows[0].Cells['MetadataCoverInfo'].Value)) { throw '整理器没有显示“跟随元数据”的实际章节封面指向。' }
        if ($undoButton.Text -ne '撤销' -or $redoButton.Text -ne '恢复') { throw '整理器缺少撤销或恢复按钮。' }
        $smokeHistoryNumber = [string]$grid.Rows[0].Cells['Number'].Value
        & $pushUndoSnapshot
        $grid.Rows[0].Cells['Number'].Value = '88'
        & $invokeUndo
        if ([string]$grid.Rows[0].Cells['Number'].Value -cne $smokeHistoryNumber) { throw '整理器撤销没有恢复上一个表格状态。' }
        & $invokeRedo
        if ([string]$grid.Rows[0].Cells['Number'].Value -cne '88') { throw '整理器恢复没有重做被撤销的操作。' }
        & $invokeUndo
        $grid.Rows[0].Cells['Batch'].Value = $true
        & $setChapterCoverRows @($grid.Rows[0]) 'first' ''
        $smokeOriginalCount = $grid.Rows.Count
        & $duplicateRowsAction
        if ($grid.Rows.Count -le $smokeOriginalCount) { throw '批量复制章节行没有生效。' }
        & $deleteRowsAction
        if ($grid.Rows.Count -ne $smokeOriginalCount) { throw '批量删除章节行没有生效。' }
        if ($grid.Rows.Count -ge 5) {
            foreach ($smokeGridRow in $grid.Rows) { $smokeGridRow.Cells['Batch'].Value = $false }
            $grid.Rows[0].Cells['Number'].Value = '1'
            $grid.Rows[1].Cells['Title'].Value = '原第2话章节名'
            $grid.Rows[2].Cells['Title'].Value = '原第3话章节名'
            $smokeMergedOriginals = @(
                [pscustomobject]@{ Number = [string]$grid.Rows[1].Cells['Number'].Value; Title = [string]$grid.Rows[1].Cells['Title'].Value; ChapterCover = [string]$grid.Rows[1].Cells['ChapterCover'].Value; Metadata = [string]$grid.Rows[1].Cells['MetadataCoverInfo'].Value },
                [pscustomobject]@{ Number = [string]$grid.Rows[2].Cells['Number'].Value; Title = [string]$grid.Rows[2].Cells['Title'].Value; ChapterCover = [string]$grid.Rows[2].Cells['ChapterCover'].Value; Metadata = [string]$grid.Rows[2].Cells['MetadataCoverInfo'].Value }
            )
            foreach ($smokeIndex in 0..2) { $grid.Rows[$smokeIndex].Cells['Batch'].Value = $true }
            & $mergeSelectedAction
            if ($grid.Rows[1].Cells['Merge'].Value -ne $true -or $grid.Rows[2].Cells['Merge'].Value -ne $true) { throw '连续多选合并没有生成正确的并入结构。' }
            for ($smokeMergedIndex = 0; $smokeMergedIndex -lt 2; $smokeMergedIndex++) {
                $smokeMergedRow = $grid.Rows[$smokeMergedIndex + 1]
                $smokeOriginal = $smokeMergedOriginals[$smokeMergedIndex]
                if ([string]$smokeMergedRow.Cells['Number'].Value -cne $smokeOriginal.Number -or [string]$smokeMergedRow.Cells['Title'].Value -cne $smokeOriginal.Title -or [string]$smokeMergedRow.Cells['ChapterCover'].Value -cne $smokeOriginal.ChapterCover -or [string]$smokeMergedRow.Cells['MetadataCoverInfo'].Value -cne $smokeOriginal.Metadata) { throw '合并后的并入行没有保留原话序、章节名或章节封面信息。' }
                if (-not $smokeMergedRow.Cells['Number'].ReadOnly -or -not $smokeMergedRow.Cells['Title'].ReadOnly -or -not $smokeMergedRow.Cells['ChapterCover'].ReadOnly) { throw '合并后的原信息没有进入灰显只读状态。' }
            }
            & $renumberRows
            $smokeLogicalRows = @($grid.Rows | Where-Object { $_.Cells['Merge'].Value -ne $true } | Sort-Object Index)
            if ([string]$smokeLogicalRows[0].Cells['Number'].Value -ne '1' -or [string]$smokeLogicalRows[1].Cells['Number'].Value -ne '2' -or [string]$smokeLogicalRows[2].Cells['Number'].Value -ne '3') { throw '合并后的逻辑章节没有按 1、2、3 重新续号。' }
            if ([string]$grid.Rows[1].Cells['Number'].Value -cne $smokeMergedOriginals[0].Number -or [string]$grid.Rows[2].Cells['Number'].Value -cne $smokeMergedOriginals[1].Number) { throw '后续编号覆盖了并入行用于辨识的原话序。' }
        }
        elseif ($grid.Rows.Count -ge 2) {
            foreach ($smokeGridRow in $grid.Rows) { $smokeGridRow.Cells['Batch'].Value = $false }
            $grid.Rows[0].Cells['Batch'].Value = $true
            $grid.Rows[$grid.Rows.Count - 1].Cells['Batch'].Value = $true
            & $mergeSelectedAction
            if ($grid.Rows[0].Cells['Merge'].Value -eq $true -or $grid.Rows[1].Cells['Merge'].Value -ne $true) { throw '非连续批量合并没有生成正确的“并入上一话”结构。' }
        }
        $script:OrganizerProgressCallback = $null
        $script:OrganizerIsRunning = $false
        $form.Dispose()
        return
    }
    try { [void]$form.ShowDialog() }
    finally {
        $script:OrganizerProgressCallback = $null
        $script:OrganizerIsRunning = $false
        $script:OrganizerBusyControlStates = @()
    }
}

try {
    $scriptPath = ''
    if (-not [string]::IsNullOrWhiteSpace($env:LOCAL_COMIC_TOOL_PATH)) {
        $scriptPath = $env:LOCAL_COMIC_TOOL_PATH
    }
    elseif ($null -ne $MyInvocation.MyCommand.PSObject.Properties['Path']) {
        $scriptPath = [string]$MyInvocation.MyCommand.Path
    }
    $scriptDirectory = if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        (Get-Location).Path
    }
    else {
        [IO.Path]::GetDirectoryName($scriptPath)
    }
    if ([string]::IsNullOrWhiteSpace($RootPath)) {
        $rootCandidate = $scriptDirectory
    }
    elseif ([IO.Path]::IsPathRooted($RootPath)) {
        $rootCandidate = $RootPath
    }
    else {
        $rootCandidate = Join-Path $scriptDirectory $RootPath
    }
    $resolvedRoot = (Resolve-Path -LiteralPath $rootCandidate).Path

    if (-not [string]::IsNullOrWhiteSpace($PlanPath)) {
        if (-not [IO.Path]::IsPathRooted($PlanPath)) {
            $PlanPath = Join-Path $resolvedRoot $PlanPath
        }
        $resolvedPlan = (Resolve-Path -LiteralPath $PlanPath).Path
        $plan = Read-OrganizerPlan -Path $resolvedPlan
        $audit = Test-OrganizerPlan -Plan $plan -LibraryRoot $resolvedRoot -IgnoreExistingOutput:$ValidateOnly
        if (-not $audit.IsValid) {
            foreach ($errorMessage in $audit.Errors) {
                Write-Host ('[错误] ' + $errorMessage) -ForegroundColor Red
            }
            exit 2
        }
        foreach ($warningMessage in $audit.Warnings) {
            Write-Host ('[警告] ' + $warningMessage) -ForegroundColor Yellow
        }
        if ($ValidateOnly) {
            Write-Host ('[完成] 方案核验通过：{0} 话，{1} 张图片。' -f $audit.ChapterCount, $audit.TotalImages) -ForegroundColor Green
            exit 0
        }
        $resultPath = Invoke-OrganizerPlan -Audit $audit -LibraryRoot $resolvedRoot
        Write-Host ('[完成] 整理并复核完成：' + $resultPath) -ForegroundColor Green
        exit 0
    }

    if ($NonInteractive) {
        throw '非交互模式必须提供 -PlanPath。'
    }
    $candidates = @(Get-CandidateFolders -LibraryRoot $resolvedRoot)
    if ($candidates.Count -eq 0) {
        throw '当前目录没有可整理的漫画文件夹。'
    }
    Show-OrganizerWindow -LibraryRoot $resolvedRoot -Candidates $candidates -SmokeTest:$UiSmokeTest
    exit 0
}
catch {
    Write-Host ('[错误] ' + $_.Exception.Message) -ForegroundColor Red
    if ($UiSmokeTest -and -not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) { Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray }
    if (-not $NonInteractive -and -not $UiSmokeTest) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '漫画整理器错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
    exit 1
}
