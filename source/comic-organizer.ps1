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
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

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
    $result = [ordered]@{ Enabled = $false; ExactMap = @{}; NormalizedMap = @{}; Warning = '' }
    $metadataPath = Join-Path $ComicPath '元数据.json'
    if (-not (Test-Path -LiteralPath $metadataPath -PathType Leaf)) { return [pscustomobject]$result }
    try {
        $metadata = [IO.File]::ReadAllText($metadataPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
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
            $entry = [pscustomobject]@{ Folder = $chapterFolder; Order = $order }
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
    if ($naturalEntries.Count -le 1) {
        foreach ($entry in $naturalEntries) {
            $entry | Add-Member -NotePropertyName OrderSource -NotePropertyValue 'Natural' -Force
            $entry | Add-Member -NotePropertyName OrderWarning -NotePropertyValue '' -Force
        }
        return @($naturalEntries)
    }
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
        [string]$SourceChapter
    )
    if (-not (Test-SimpleFolderName $SourceFolder)) {
        throw ('来源漫画文件夹名称不合法：' + $SourceFolder)
    }
    $comicPath = Join-Path $LibraryRoot $SourceFolder
    if (-not (Test-Path -LiteralPath $comicPath -PathType Container)) {
        throw ('来源漫画文件夹不存在：' + $SourceFolder)
    }
    $entries = @(Get-SourceChapterEntries -ComicPath $comicPath)
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
    }
}

function Get-CandidateFolders {
    param([string]$LibraryRoot)
    $candidates = @()
    foreach ($directory in @(Get-ChildItem -LiteralPath $LibraryRoot -Directory | Sort-Object Name)) {
        if ($directory.Name -eq $script:OutputFolderName) { continue }
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
        [switch]$IgnoreExistingOutput
    )

    $errors = New-Object 'System.Collections.Generic.List[string]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    $omissionWarnings = New-Object 'System.Collections.Generic.List[string]'
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
    if ($planChapters.Count -eq 0) {
        $errors.Add('方案中没有章节。')
    }

    $resolvedChapters = @()
    $sourceCache = @{}
    $usageBySource = @{}
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
        $mergeWithPrevious = [bool](Get-ObjectProperty -Object $chapter -Name 'mergeWithPrevious' -Default $false)
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
                $sourceCache[$sourceKey] = Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $sourceFolder -SourceChapter $sourceChapter
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

        if (-not $usageBySource.ContainsKey($sourceKey)) {
            $usageBySource[$sourceKey] = @{}
        }
        $usage = $usageBySource[$sourceKey]
        $hasOverlap = $false
        for ($imageNumber = $start; $imageNumber -le $end; $imageNumber++) {
            if ($usage.ContainsKey([string]$imageNumber)) {
                $errors.Add(('图片被重复使用：{0}\{1}\{2:D4}（{3} 与 {4} 重叠）' -f $sourceFolder, $sourceChapter, $imageNumber, $usage[[string]$imageNumber], $chapterLabel))
                $hasOverlap = $true
            }
            else {
                $usage[[string]$imageNumber] = $chapterLabel
            }
        }
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

    foreach ($sourceFolder in $selectedSourceFolders) {
        $comicPath = Join-Path $LibraryRoot $sourceFolder
        try {
            $sourceChapterEntries = @(Get-SourceChapterEntries -ComicPath $comicPath)
        }
        catch {
            $errors.Add(($sourceFolder + '：' + $_.Exception.Message))
            continue
        }
        foreach ($chapterDirectory in $sourceChapterEntries) {
            $sourceKey = $sourceFolder + '|' + $chapterDirectory.Name
            try {
                if (-not $sourceCache.ContainsKey($sourceKey)) {
                    $sourceCache[$sourceKey] = Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $sourceFolder -SourceChapter $chapterDirectory.Name
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
    if (-not (Test-SimpleFolderName $coverSource)) {
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
                $coverChapterDirectories = @(Get-SourceChapterEntries -ComicPath $coverComicPath)
                if ($coverChapterDirectories.Count -eq 0) {
                    throw '没有可识别的章节文件夹。'
                }
                $coverChapterInfo = Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $coverSource -SourceChapter $coverChapterDirectories[0].Name
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
    return [pscustomobject]@{
        IsValid = ($errors.Count -eq 0)
        Errors = @($errors)
        Warnings = @($warnings)
        OmissionWarnings = @($omissionWarnings)
        OutputName = $outputName
        OutputBase = $outputBase
        OutputPath = $outputPath
        CoverSource = $coverSource
        CoverPath = $coverPath
        CoverOutputName = $coverOutputName
        CoverIsAutomatic = $coverIsAutomatic
        Description = $description
        DescriptionSource = $descriptionSource
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
        }
    })
    $organizerInfo = [pscustomobject][ordered]@{
        schemaVersion = 4
        generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        sourceFolders = @($Audit.SourceFolders)
        chapterCount = $Audit.ChapterCount
        totalImages = $Audit.TotalImages
        coverFile = $Audit.CoverOutputName
        coverAutomatic = $Audit.CoverIsAutomatic
        descriptionSource = $Audit.DescriptionSource
    }
    $metadata | Add-Member -NotePropertyName name -NotePropertyValue $Audit.OutputName -Force
    $metadata | Add-Member -NotePropertyName description -NotePropertyValue $Audit.Description -Force
    $metadata | Add-Member -NotePropertyName chapterInfos -NotePropertyValue $chapterInfos -Force
    $metadata | Add-Member -NotePropertyName organizer -NotePropertyValue $organizerInfo -Force
    return $metadata
}

function Test-NormalizedOutput {
    param(
        [string]$OutputPath,
        [object]$Audit
    )
    $errors = New-Object 'System.Collections.Generic.List[string]'
    $outputCoverPath = Join-Path $OutputPath $Audit.CoverOutputName
    if (-not (Test-Path -LiteralPath $outputCoverPath -PathType Leaf)) {
        $errors.Add(('输出缺少封面：' + $Audit.CoverOutputName))
    }
    elseif ((Get-Item -LiteralPath $outputCoverPath).Length -eq 0) {
        $errors.Add(('输出封面是空文件：' + $Audit.CoverOutputName))
    }
    foreach ($chapter in $Audit.Chapters) {
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
    }
    return @($errors)
}

function Invoke-OrganizerPlan {
    param(
        [object]$Audit,
        [string]$LibraryRoot
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
        Copy-Item -LiteralPath $Audit.CoverPath -Destination (Join-Path $temporaryPath $Audit.CoverOutputName)

        foreach ($chapter in $Audit.Chapters) {
            $folderName = $chapter.FolderName
            $chapterPath = Join-Path $temporaryPath $folderName
            [void](New-Item -ItemType Directory -Path $chapterPath)
            $outputNumber = 0
            foreach ($image in $chapter.Images) {
                $outputNumber++
                $destinationName = $outputNumber.ToString('D4') + $image.Extension.ToLowerInvariant()
                Copy-Item -LiteralPath $image.FullName -Destination (Join-Path $chapterPath $destinationName)
            }
        }

        $metadata = New-OutputMetadata -Audit $Audit -LibraryRoot $LibraryRoot
        Write-JsonFile -Path (Join-Path $temporaryPath '元数据.json') -Value $metadata
        Write-JsonFile -Path (Join-Path $temporaryPath '整理方案.json') -Value $Audit.Plan

        $verifyErrors = @(Test-NormalizedOutput -OutputPath $temporaryPath -Audit $Audit)
        if ($verifyErrors.Count -gt 0) {
            throw ('输出复核失败：' + ($verifyErrors -join '；'))
        }
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
    $form.Size = New-Object System.Drawing.Size(1280, 800)
    $form.MinimumSize = New-Object System.Drawing.Size(1080, 680)
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
    $outputName.Size = New-Object System.Drawing.Size(565, 27)
    $outputName.Anchor = 'Top,Left,Right'
    $form.Controls.Add($outputName)

    $descriptionButton = New-Object System.Windows.Forms.Button
    $descriptionButton.Text = '编辑简介（未填写）'
    $descriptionButton.Location = New-Object System.Drawing.Point(1083, 75)
    $descriptionButton.Size = New-Object System.Drawing.Size(162, 32)
    $descriptionButton.Anchor = 'Top,Right'
    $descriptionButton.BackColor = [System.Drawing.Color]::FromArgb(238, 244, 250)
    $form.Controls.Add($descriptionButton)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(405, 118)
    $grid.Size = New-Object System.Drawing.Size(840, 440)
    $grid.Anchor = 'Top,Bottom,Left,Right'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.MultiSelect = $false
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
    $grid.Columns['Number'].Width = 85
    $grid.Columns['Number'].ToolTipText = '可填 4.5、1.01，也可直接填“特典话”“番外篇”“插画集”；实际阅读位置由行顺序决定'
    $grid.Columns['Title'].Width = 180
    $grid.Columns['Title'].ToolTipText = '每一行都可单独命名；输出示例：第4.5话 特典篇'
    $grid.Columns['SourceFolder'].Width = 190
    $grid.Columns['SourceChapter'].Width = 100
    $grid.Columns['Start'].Width = 65
    $grid.Columns['End'].Width = 65
    $grid.Columns['Total'].Width = 55
    $grid.Columns['Merge'].Width = 82
    $grid.Columns['Cover'].Width = 55
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

    $moveUp = New-Object System.Windows.Forms.Button
    $moveUp.Text = '上移'
    $moveUp.Location = New-Object System.Drawing.Point(814, 572)
    $moveUp.Size = New-Object System.Drawing.Size(70, 34)
    $moveUp.Anchor = 'Bottom,Left'
    $form.Controls.Add($moveUp)

    $moveDown = New-Object System.Windows.Forms.Button
    $moveDown.Text = '下移'
    $moveDown.Location = New-Object System.Drawing.Point(892, 572)
    $moveDown.Size = New-Object System.Drawing.Size(70, 34)
    $moveDown.Anchor = 'Bottom,Left'
    $form.Controls.Add($moveDown)

    $autoNumber = New-Object System.Windows.Forms.Button
    $autoNumber.Text = '从选中行后续编号'
    $autoNumber.Location = New-Object System.Drawing.Point(970, 572)
    $autoNumber.Size = New-Object System.Drawing.Size(135, 34)
    $autoNumber.Anchor = 'Bottom,Left'
    $form.Controls.Add($autoNumber)

    $validateButton = New-Object System.Windows.Forms.Button
    $validateButton.Text = '核验方案'
    $validateButton.Location = New-Object System.Drawing.Point(405, 620)
    $validateButton.Size = New-Object System.Drawing.Size(110, 38)
    $validateButton.Anchor = 'Bottom,Left'
    $form.Controls.Add($validateButton)

    $savePlan = New-Object System.Windows.Forms.Button
    $savePlan.Text = '保存方案'
    $savePlan.Location = New-Object System.Drawing.Point(523, 620)
    $savePlan.Size = New-Object System.Drawing.Size(110, 38)
    $savePlan.Anchor = 'Bottom,Left'
    $form.Controls.Add($savePlan)

    $loadPlan = New-Object System.Windows.Forms.Button
    $loadPlan.Text = '加载方案'
    $loadPlan.Location = New-Object System.Drawing.Point(641, 620)
    $loadPlan.Size = New-Object System.Drawing.Size(110, 38)
    $loadPlan.Anchor = 'Bottom,Left'
    $form.Controls.Add($loadPlan)

    $helpButton = New-Object System.Windows.Forms.Button
    $helpButton.Text = '使用说明'
    $helpButton.Location = New-Object System.Drawing.Point(759, 620)
    $helpButton.Size = New-Object System.Drawing.Size(110, 38)
    $helpButton.Anchor = 'Bottom,Left'
    $form.Controls.Add($helpButton)

    $generateButton = New-Object System.Windows.Forms.Button
    $generateButton.Text = '开始整理'
    $generateButton.Location = New-Object System.Drawing.Point(1105, 620)
    $generateButton.Size = New-Object System.Drawing.Size(140, 38)
    $generateButton.Anchor = 'Bottom,Right'
    $generateButton.BackColor = [System.Drawing.Color]::FromArgb(35, 105, 160)
    $generateButton.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($generateButton)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = '等待载入来源文件夹。'
    $status.AutoSize = $false
    $status.Location = New-Object System.Drawing.Point(18, 680)
    $status.Size = New-Object System.Drawing.Size(1227, 55)
    $status.Anchor = 'Bottom,Left,Right'
    $status.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($status)

    $script:lastDefaultOutput = ''
    $script:OrganizerDescription = ''
    $script:OrganizerDescriptionSource = ''
    $script:OrganizerDescriptionLocked = $false
    $script:OrganizerLoadedOnce = $false

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

    $outputName.add_SelectedIndexChanged({ & $syncDefaultDescriptionFromOutputName })

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

    $renumberRows = {
        if ($grid.Rows.Count -eq 0) { return }
        if ($null -eq $grid.CurrentRow) {
            for ($index = 0; $index -lt $grid.Rows.Count; $index++) {
                $grid.Rows[$index].Cells['Number'].Value = $index + 1
            }
            $status.Text = '没有选中锚点，已从第1行开始按整数编号。'
            return
        }
        $anchorIndex = $grid.CurrentRow.Index
        $anchorNumber = ConvertTo-ChapterNumberInfo -Value $grid.CurrentRow.Cells['Number'].Value
        if ($null -eq $anchorNumber) {
            & $showMessage '选中行的话序无效；请先填写正整数或小数，例如 3 或 4.5。' '无法继续编号' ([System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $nextNumber = [int][decimal]::Floor($anchorNumber.Value) + 1
        for ($index = $anchorIndex + 1; $index -lt $grid.Rows.Count; $index++) {
            $grid.Rows[$index].Cells['Number'].Value = $nextNumber
            $nextNumber++
        }
        $status.Text = ('已保留前 {0} 行，从选中行的下一行开始编号。' -f ($anchorIndex + 1))
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
        if ($grid.IsCurrentCellDirty -and $null -ne $grid.CurrentCell -and $grid.CurrentCell.OwningColumn.Name -eq 'Cover') {
            [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })
    $grid.add_CellValueChanged({
        param($sender, $eventArgs)
        if ($coverSelectionState.Changing -or $eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -lt 0) { return }
        if ($grid.Columns[$eventArgs.ColumnIndex].Name -ne 'Cover') { return }
        $row = $grid.Rows[$eventArgs.RowIndex]
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
            }
        }
        if ([string]::IsNullOrWhiteSpace($selectedCoverSource) -and $grid.Rows.Count -gt 0) {
            $selectedCoverSource = [string]$grid.Rows[0].Cells['SourceFolder'].Value
        }
        return [pscustomobject][ordered]@{
            schemaVersion = 5
            outputName = $outputName.Text.Trim()
            coverSource = $selectedCoverSource
            descriptionSource = $script:OrganizerDescriptionSource
            description = $script:OrganizerDescription
            descriptionLocked = $script:OrganizerDescriptionLocked
            selectedSourceFolders = @($sourceList.CheckedItems | ForEach-Object { [string]$_ })
            chapters = $chapters
        }
    }

    $setGridFromPlan = {
        param([object]$Plan)
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
        foreach ($chapter in @((Get-ObjectProperty -Object $Plan -Name 'chapters' -Default @()))) {
            $sourceFolder = [string](Get-ObjectProperty -Object $chapter -Name 'sourceFolder' -Default '')
            $sourceChapter = [string](Get-ObjectProperty -Object $chapter -Name 'sourceChapter' -Default '')
            $total = ''
            try {
                $total = (Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $sourceFolder -SourceChapter $sourceChapter).Count
            }
            catch {}
            [void]$grid.Rows.Add(
                [string](Get-ObjectProperty -Object $chapter -Name 'number' -Default ''),
                [string](Get-ObjectProperty -Object $chapter -Name 'title' -Default ''),
                $sourceFolder,
                $sourceChapter,
                [string](Get-ObjectProperty -Object $chapter -Name 'start' -Default ''),
                [string](Get-ObjectProperty -Object $chapter -Name 'end' -Default ''),
                [string]$total,
                [bool](Get-ObjectProperty -Object $chapter -Name 'mergeWithPrevious' -Default $false),
                $false
            )
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
                    [void]$grid.Rows.Add($initialFields.Number, $initialFields.Title, $name, $chapterDirectory.Name, 1, $imageCount, $imageCount, $false, ($grid.Rows.Count -eq 0))
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
        $currentOutputText = $outputName.Text
        $loadResult = & $addSourceRows $newNames $false
        if ($null -eq $loadResult) { return }
        $allLoadedNames = @(& $getLoadedSourceNames)
        & $refreshOutputNameChoices $allLoadedNames $currentOutputText
        $status.Text = ('已追加 {0} 个新来源、{1} 行章节；原有 {2} 行编辑未重置。' -f $newNames.Count, $loadResult.AddedRows, ($grid.Rows.Count - $loadResult.AddedRows))
        & $showOrderWarnings $loadResult
    })

    $rescanSources.add_Click({
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
            foreach ($row in $grid.Rows) {
                $sourceFolder = [string]$row.Cells['SourceFolder'].Value
                $sourceChapter = [string]$row.Cells['SourceChapter'].Value
                $oldTotal = 0
                [void][int]::TryParse([string]$row.Cells['Total'].Value, [ref]$oldTotal)
                try {
                    $newTotal = (Get-SourceChapterInfo -LibraryRoot $LibraryRoot -SourceFolder $sourceFolder -SourceChapter $sourceChapter).Count
                    $endValue = 0
                    $endWasTotal = [int]::TryParse([string]$row.Cells['End'].Value, [ref]$endValue) -and $oldTotal -gt 0 -and $endValue -eq $oldTotal
                    if ($endWasTotal -or [string]::IsNullOrWhiteSpace([string]$row.Cells['End'].Value)) { $row.Cells['End'].Value = $newTotal }
                    $row.Cells['Total'].Value = $newTotal
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

    $duplicateRow.add_Click({
        if ($null -eq $grid.CurrentRow) { return }
        $sourceRow = $grid.CurrentRow
        $index = $sourceRow.Index + 1
        $values = @(
            '',
            [string]$sourceRow.Cells['Title'].Value,
            [string]$sourceRow.Cells['SourceFolder'].Value,
            [string]$sourceRow.Cells['SourceChapter'].Value,
            [string]$sourceRow.Cells['Start'].Value,
            [string]$sourceRow.Cells['End'].Value,
            [string]$sourceRow.Cells['Total'].Value,
            $false,
            $false
        )
        $grid.Rows.Insert($index, $values)
        $grid.CurrentCell = $grid.Rows[$index].Cells['Number']
        $status.Text = '已复制一行；请填写新行话序/特殊标签，或选中数字锚点后点击“从选中行后续编号”。'
    })

    $deleteRow.add_Click({
        if ($null -eq $grid.CurrentRow) { return }
        $deletedIndex = $grid.CurrentRow.Index
        $deletedCover = ($grid.CurrentRow.Cells['Cover'].Value -eq $true)
        $grid.Rows.RemoveAt($deletedIndex)
        if ($deletedCover -and $grid.Rows.Count -gt 0) {
            & $setCoverRow ([Math]::Min($deletedIndex, $grid.Rows.Count - 1))
        }
        $status.Text = '已删除一行；现有话序保持不变。'
    })

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
            $ranges += ,@($startValue, $endValue)
        }
        $insertIndex = $row.Index
        $sourceFolderValue = [string]$row.Cells['SourceFolder'].Value
        $sourceChapterValue = [string]$row.Cells['SourceChapter'].Value
        $totalValue = [string]$row.Cells['Total'].Value
        $originalNumber = [string]$row.Cells['Number'].Value
        $wasMerged = ($row.Cells['Merge'].Value -eq $true)
        $wasCover = ($row.Cells['Cover'].Value -eq $true)
        $grid.Rows.RemoveAt($insertIndex)
        for ($offset = 0; $offset -lt $ranges.Count; $offset++) {
            $range = $ranges[$offset]
            $grid.Rows.Insert(
                $insertIndex + $offset,
                @($(if ($offset -eq 0) { $originalNumber } else { '' }), '', $sourceFolderValue, $sourceChapterValue, $range[0], $range[1], $totalValue, ($wasMerged -and $offset -eq 0), ($wasCover -and $offset -eq 0))
            )
        }
        if ($wasCover) { & $setCoverRow $insertIndex }
        $grid.CurrentCell = $grid.Rows[$insertIndex].Cells['Number']
        $status.Text = ('已将一行拆成 {0} 话；请填写新增行话序/特殊标签（如 4.5、特典话），或选中数字锚点后点击“从选中行后续编号”。' -f $ranges.Count)
    })

    $moveRow = {
        param([int]$Delta)
        if ($null -eq $grid.CurrentRow) { return }
        $oldIndex = $grid.CurrentRow.Index
        $newIndex = $oldIndex + $Delta
        if ($newIndex -lt 0 -or $newIndex -ge $grid.Rows.Count) { return }
        $values = @()
        foreach ($cell in $grid.Rows[$oldIndex].Cells) { $values += $cell.Value }
        $grid.Rows.RemoveAt($oldIndex)
        $grid.Rows.Insert($newIndex, $values)
        $grid.CurrentCell = $grid.Rows[$newIndex].Cells['Number']
        $status.Text = '已移动一行；表格从上到下就是实际阅读顺序，话序标签保持不变。'
    }
    $moveUp.add_Click({ & $moveRow -1 })
    $moveDown.add_Click({ & $moveRow 1 })
    $autoNumber.add_Click({ & $renumberRows })

    $helpButton.add_Click({
        $helpText = @'
1. 首次勾选来源后点击“载入所选文件夹”；载入后该按钮会变为“重新载入”，使用它会重建表格并重置现有编辑。
2. 编辑中途新增来源时，先“重新扫描来源文件夹库”，再勾选新来源并点击“载入新增文件夹”，原有行不会重置。
3. 输出漫画名称可手动输入，也可用右侧下拉箭头直接选择已载入漫画名称。
4. 简介默认跟随输出漫画名称所选来源；在简介窗口按“确定”后即固定，不再随名称来源变化。
5. 来源若有元数据.json 且 chapterInfos 完整有效，会按其中的 order 排列；无配置或无法完整匹配时按名称自然排序并提示原因。
6. 根目录若为 P01_001、P02_001 等格式，会按前缀自动生成多行章节；无法识别页码时按名称自然排序。
7. 表格从上到下就是阅读顺序，可修改话序、章节名、范围并上下移动；非数字话序会保留原特殊名称。
8. 若要把多个文件夹或分组接成同一话，把来源行排在一起，并从第二行起勾选“并入上一话”。
9. “按范围拆分选中行”可把总集的一行拆成多话。
10. 原文件不会修改；结果输出到“整理完成”，正文统一重命名为 0001、0002……。
11. 纯数字或复合页码会检查重复和缺号；无法识别的名称排序模式会明确警告无法判断缺图。
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
            try {
                & $setGridFromPlan (Read-OrganizerPlan -Path $dialog.FileName)
            }
            catch {
                & $showMessage $_.Exception.Message '加载失败' ([System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    $generateButton.add_Click({
        try {
            $plan = & $getPlanFromGrid
            $audit = Test-OrganizerPlan -Plan $plan -LibraryRoot $LibraryRoot
            if (-not $audit.IsValid) {
                & $showMessage (($audit.Errors | Select-Object -First 30) -join "`r`n") '方案未通过' ([System.Windows.Forms.MessageBoxIcon]::Error)
                return
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
            $form.UseWaitCursor = $true
            [System.Windows.Forms.Application]::DoEvents()
            $resultPath = Invoke-OrganizerPlan -Audit $audit -LibraryRoot $LibraryRoot
            $form.UseWaitCursor = $false
            $status.Text = ('整理完成：' + $resultPath)
            & $showMessage ("整理并复核完成。`r`n`r`n$resultPath`r`n`r`n确认无误后，可手动复制到漫画大文件夹。") '整理完成' ([System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            $form.UseWaitCursor = $false
            & $showMessage $_.Exception.Message '整理失败' ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    if ($SmokeTest) {
        $form.Dispose()
        return
    }
    [void]$form.ShowDialog()
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
    if (-not $NonInteractive -and -not $UiSmokeTest) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '漫画整理器错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
    exit 1
}
