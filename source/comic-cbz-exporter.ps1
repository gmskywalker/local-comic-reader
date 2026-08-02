[CmdletBinding()]
param(
    [string]$RootPath = '',
    [string]$OutputPath = '',
    [string]$ComicName = '',
    [ValidateSet('PerChapter', 'SingleBook', 'Epub')]
    [string]$Mode = 'PerChapter',
    [switch]$IncludeCover,
    [switch]$NonInteractive,
    [switch]$ValidateOnly,
    [switch]$ForceIssues,
    [switch]$UiSmokeTest,
    [switch]$SkipOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ImageExtensions = @('.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.avif')
$script:DefaultOutputFolderName = 'CBZ导出'
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Get-ObjectProperty {
    param([object]$Object, [string]$Name, [object]$DefaultValue = $null)
    if ($null -eq $Object) { return $DefaultValue }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $DefaultValue }
    return $property.Value
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
        if ($orderedNumbers.Count -gt 0 -and $orderedNumbers[0] -ne 1) {
            $issues.Add(('{0}：第一张图片通常应为 0001，实际编号为 {1}。' -f $Context, $orderedNumbers[0]))
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
    foreach ($child in @(Get-ChildItem -LiteralPath $ComicPath -Directory -ErrorAction SilentlyContinue)) { $queue.Enqueue($child) }
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
        foreach ($child in @(Get-ChildItem -LiteralPath $directory.FullName -Directory -ErrorAction SilentlyContinue)) { $queue.Enqueue($child) }
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
        $result.Title = [string](Get-ObjectProperty -Object $metadata -Name 'name' -DefaultValue '')
        $authors = @(Get-ObjectProperty -Object $metadata -Name 'author' -DefaultValue @()) | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $result.Author = $authors -join '、'
        $result.Description = [string](Get-ObjectProperty -Object $metadata -Name 'description' -DefaultValue '')
        $chapterInfos = @(Get-ObjectProperty -Object $metadata -Name 'chapterInfos' -DefaultValue @())
        if ($chapterInfos.Count -eq 0) {
            $result.Warning = '整理器元数据没有 chapterInfos，已改用名称自然排序。'
            return [pscustomobject]$result
        }
        $infos = @()
        $usedFolders = @{}
        $usedOrders = @{}
        foreach ($item in $chapterInfos) {
            $folder = [string](Get-ObjectProperty $item 'chapterFolder' '')
            if ([string]::IsNullOrWhiteSpace($folder)) { $folder = [string](Get-ObjectProperty $item 'chapterTitle' '') }
            $order = [double]0
            $validOrder = [double]::TryParse([string](Get-ObjectProperty $item 'order' 0), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$order)
            $orderKey = $order.ToString('R', [Globalization.CultureInfo]::InvariantCulture)
            if ([string]::IsNullOrWhiteSpace($folder) -or -not $validOrder -or $order -le 0 -or $usedFolders.ContainsKey($folder) -or $usedOrders.ContainsKey($orderKey)) {
                $result.Warning = '整理器阅读顺序含有空名称、重复名称或重复顺序，已改用名称自然排序。'
                return [pscustomobject]$result
            }
            $label = [string](Get-ObjectProperty $item 'displayLabel' '')
            if ([string]::IsNullOrWhiteSpace($label)) { $label = $folder }
            $infos += [pscustomobject]@{ Folder = $folder; Order = $order; Label = $label }
            $usedFolders[$folder] = $true
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
    param([System.IO.DirectoryInfo]$ComicDirectory)
    $issues = New-Object 'System.Collections.Generic.List[string]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    $cover = Get-CoverFile -ComicPath $ComicDirectory.FullName
    $rootImages = @(Get-RootBodyImages -ComicPath $ComicDirectory.FullName)

    if ($null -eq $cover) {
        $zeroCandidates = @($rootImages | Where-Object { $_.BaseName -match '^0+$' })
        if ($zeroCandidates.Count -eq 1 -and $rootImages.Count -gt 1) {
            $cover = $zeroCandidates[0]
            $rootImages = @($rootImages | Where-Object { $_.FullName -ine $cover.FullName })
            $warnings.Add(('已把根目录的零号图片作为封面：' + $cover.Name))
        }
    }

    $directoryScan = Get-ImageChapterDirectories -ComicPath $ComicDirectory.FullName
    foreach ($message in $directoryScan.Warnings) { $warnings.Add($message) }
    $chapters = @()

    if ($rootImages.Count -gt 0) {
        $rootLayout = Get-RootCompositeGroups -Files $rootImages
        if ($rootLayout.Recognized) {
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
            $warnings.Add(('根目录图片已按文件名前缀识别为 {0} 个章节组。' -f $rootLayout.Groups.Count))
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
    $queue = New-Object 'System.Collections.Generic.Queue[System.IO.DirectoryInfo]'
    $queue.Enqueue($Directory)
    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
        $found = @(Get-ChildItem -LiteralPath $current.FullName -File -ErrorAction SilentlyContinue | Where-Object { Test-ImageFile -File $_ } | Select-Object -First 1)
        if ($found.Count -gt 0) { return $true }
        foreach ($child in @(Get-ChildItem -LiteralPath $current.FullName -Directory -ErrorAction SilentlyContinue)) { $queue.Enqueue($child) }
    }
    return $false
}

function Get-CandidateComics {
    param([string]$LibraryRoot, [string]$ResolvedOutputPath)
    $result = @()
    foreach ($directory in @(Get-ChildItem -LiteralPath $LibraryRoot -Directory -ErrorAction Stop | Sort-Object { Get-NaturalNameSortKey $_.Name }, Name)) {
        if (-not [string]::IsNullOrWhiteSpace($ResolvedOutputPath) -and $directory.FullName -ieq $ResolvedOutputPath) { continue }
        if ($directory.Name -in @('.git', 'CBZ导出')) { continue }
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

function Test-PathInside {
    param([string]$ChildPath, [string]$ParentPath)
    $child = [IO.Path]::GetFullPath($ChildPath).TrimEnd('\')
    $parent = [IO.Path]::GetFullPath($ParentPath).TrimEnd('\')
    return $child.Equals($parent, [StringComparison]::OrdinalIgnoreCase) -or $child.StartsWith($parent + '\', [StringComparison]::OrdinalIgnoreCase)
}

function New-CbzArchive {
    param(
        [string]$DestinationPath,
        [System.IO.FileInfo[]]$Images,
        [System.IO.FileInfo]$Cover = $null,
        [int]$Digits = 6
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
                    $coverName = ('{0}-cover{1}' -f ('0' * $Digits), $Cover.Extension.ToLowerInvariant())
                    $entry = $archive.CreateEntry($coverName, [IO.Compression.CompressionLevel]::Optimal)
                    $source = [IO.File]::OpenRead($Cover.FullName)
                    $target = $entry.Open()
                    try { $source.CopyTo($target) } finally { $target.Dispose(); $source.Dispose() }
                    $expected[$coverName] = $Cover.Length
                }
                foreach ($image in @($Images)) {
                    if ($null -ne $Cover -and $image.FullName -ieq $Cover.FullName) { continue }
                    $page++
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
        [System.IO.FileInfo]$Cover = $null
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
        [bool]$UseCover
    )
    $safeComic = ConvertTo-SafeFileName -Name $Plan.Name -MaxLength 82
    $results = @()
    if ($ExportMode -eq 'Epub') {
        $target = Join-Path $ResolvedOutputPath ($safeComic + '.epub')
        $cover = if ($UseCover) { $Plan.Cover } else { $null }
        $results += New-EpubArchive -DestinationPath $target -Plan $Plan -Cover $cover
        return @($results)
    }
    if ($ExportMode -eq 'SingleBook') {
        $allImages = @()
        foreach ($chapter in $Plan.Chapters) { $allImages += @($chapter.Images) }
        $target = Join-Path $ResolvedOutputPath ($safeComic + '.cbz')
        $cover = if ($UseCover) { $Plan.Cover } else { $null }
        $results += New-CbzArchive -DestinationPath $target -Images $allImages -Cover $cover -Digits 6
        return @($results)
    }

    $comicOutput = Join-Path $ResolvedOutputPath $safeComic
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
        $label = ConvertTo-SafeFileName -Name $chapter.Label -MaxLength 92
        $baseName = (($index + 1).ToString(('D' + $width))) + ' - ' + $label
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
        $results += New-CbzArchive -DestinationPath $target -Images $chapter.Images -Cover $cover -Digits 6
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
    return $dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::Yes
}

function Show-ExporterWindow {
    param([string]$LibraryRoot, [string]$InitialOutput, [switch]$SmokeTest)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    $form.Text = '本地漫画 CBZ / EPUB 导出器'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object Drawing.Size(1040, 760)
    $form.MinimumSize = New-Object Drawing.Size(880, 650)
    $form.Font = New-Object Drawing.Font('Microsoft YaHei UI', 10)

    $header = New-Object Windows.Forms.Label
    $header.Text = '漫画 CBZ / EPUB 导出器'
    $header.Font = New-Object Drawing.Font('Microsoft YaHei UI', 19, [Drawing.FontStyle]::Bold)
    $header.AutoSize = $true
    $header.Location = New-Object Drawing.Point(24, 18)
    $sub = New-Object Windows.Forms.Label
    $sub.Text = 'JSON 只用于读取章节顺序；源 HTML 不会复制。EPUB 会生成真正可点击的章节目录。'
    $sub.ForeColor = [Drawing.Color]::DimGray
    $sub.AutoSize = $true
    $sub.Location = New-Object Drawing.Point(27, 60)

    $listLabel = New-Object Windows.Forms.Label
    $listLabel.Text = '选择漫画：'
    $listLabel.AutoSize = $true
    $listLabel.Location = New-Object Drawing.Point(26, 96)
    $list = New-Object Windows.Forms.CheckedListBox
    $list.CheckOnClick = $true
    $list.HorizontalScrollbar = $true
    $list.Anchor = 'Top,Bottom,Left,Right'
    $list.Location = New-Object Drawing.Point(28, 124)
    $list.Size = New-Object Drawing.Size(970, 354)

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

    $modeGroup = New-Object Windows.Forms.GroupBox
    $modeGroup.Text = '导出方式'
    $modeGroup.Anchor = 'Bottom,Left'
    $modeGroup.Location = New-Object Drawing.Point(28, 528)
    $modeGroup.Size = New-Object Drawing.Size(430, 112)
    $perChapter = New-Object Windows.Forms.RadioButton
    $perChapter.Text = '每话一个 CBZ'
    $perChapter.Checked = $false
    $perChapter.AutoSize = $true
    $perChapter.Location = New-Object Drawing.Point(18, 26)
    $singleBook = New-Object Windows.Forms.RadioButton
    $singleBook.Text = '整部漫画一个 CBZ（无章节目录）'
    $singleBook.AutoSize = $true
    $singleBook.Location = New-Object Drawing.Point(18, 56)
    $epubBook = New-Object Windows.Forms.RadioButton
    $epubBook.Text = '整部 EPUB（ReadEra 可点击目录，推荐）'
    $epubBook.Checked = $true
    $epubBook.AutoSize = $true
    $epubBook.Location = New-Object Drawing.Point(18, 84)
    [void]$modeGroup.Controls.Add($perChapter)
    [void]$modeGroup.Controls.Add($singleBook)
    [void]$modeGroup.Controls.Add($epubBook)

    $optionsGroup = New-Object Windows.Forms.GroupBox
    $optionsGroup.Text = '选项'
    $optionsGroup.Anchor = 'Bottom,Left'
    $optionsGroup.Location = New-Object Drawing.Point(470, 528)
    $optionsGroup.Size = New-Object Drawing.Size(310, 112)
    $coverCheck = New-Object Windows.Forms.CheckBox
    $coverCheck.Text = '把总封面放在第一话 / 整本开头'
    $coverCheck.Checked = $true
    $coverCheck.AutoSize = $true
    $coverCheck.Location = New-Object Drawing.Point(16, 27)
    $openCheck = New-Object Windows.Forms.CheckBox
    $openCheck.Text = '完成后打开输出文件夹'
    $openCheck.Checked = $true
    $openCheck.AutoSize = $true
    $openCheck.Location = New-Object Drawing.Point(16, 57)
    [void]$optionsGroup.Controls.Add($coverCheck)
    [void]$optionsGroup.Controls.Add($openCheck)

    $outputLabel = New-Object Windows.Forms.Label
    $outputLabel.Text = '输出目录：'
    $outputLabel.Anchor = 'Bottom,Left'
    $outputLabel.AutoSize = $true
    $outputLabel.Location = New-Object Drawing.Point(28, 646)
    $outputBox = New-Object Windows.Forms.TextBox
    $outputBox.Anchor = 'Bottom,Left,Right'
    $outputBox.Location = New-Object Drawing.Point(112, 642)
    $outputBox.Size = New-Object Drawing.Size(700, 30)
    $outputBox.Text = $InitialOutput
    $browse = New-Object Windows.Forms.Button
    $browse.Text = '浏览…'
    $browse.Anchor = 'Bottom,Right'
    $browse.Location = New-Object Drawing.Point(822, 640)
    $browse.Size = New-Object Drawing.Size(80, 34)
    $export = New-Object Windows.Forms.Button
    $export.Text = '开始导出'
    $export.Anchor = 'Bottom,Right'
    $export.BackColor = [Drawing.Color]::FromArgb(45, 118, 174)
    $export.ForeColor = [Drawing.Color]::White
    $export.Location = New-Object Drawing.Point(910, 640)
    $export.Size = New-Object Drawing.Size(90, 34)

    $status = New-Object Windows.Forms.Label
    $status.Anchor = 'Bottom,Left,Right'
    $status.AutoEllipsis = $true
    $status.Location = New-Object Drawing.Point(28, 686)
    $status.Size = New-Object Drawing.Size(970, 24)
    $status.ForeColor = [Drawing.Color]::DimGray

    foreach ($control in @($header, $sub, $listLabel, $list, $selectAll, $selectNone, $refresh, $modeGroup, $optionsGroup, $outputLabel, $outputBox, $browse, $export, $status)) {
        [void]$form.Controls.Add($control)
    }

    $script:CandidateItems = @()
    $loadCandidates = {
        $form.Cursor = [Windows.Forms.Cursors]::WaitCursor
        $status.Text = '正在扫描漫画文件夹…'
        [Windows.Forms.Application]::DoEvents()
        try {
            $resolvedOutput = Resolve-OutputPath -LibraryRoot $LibraryRoot -Candidate $outputBox.Text
            $script:CandidateItems = @(Get-CandidateComics -LibraryRoot $LibraryRoot -ResolvedOutputPath $resolvedOutput)
            $list.Items.Clear()
            foreach ($candidate in $script:CandidateItems) {
                $hasMetadata = Test-Path -LiteralPath (Join-Path $candidate.FullName '元数据.json') -PathType Leaf
                $suffix = if ($hasMetadata) { '  [有元数据]' } else { '' }
                [void]$list.Items.Add($candidate.Name + $suffix, $false)
            }
            $status.Text = ('找到 {0} 部可导出漫画；未默认勾选，避免误导出。' -f $script:CandidateItems.Count)
        }
        finally { $form.Cursor = [Windows.Forms.Cursors]::Default }
    }

    $selectAll.Add_Click({ for ($index = 0; $index -lt $list.Items.Count; $index++) { $list.SetItemChecked($index, $true) } })
    $selectNone.Add_Click({ for ($index = 0; $index -lt $list.Items.Count; $index++) { $list.SetItemChecked($index, $false) } })
    $refresh.Add_Click($loadCandidates)
    $browse.Add_Click({
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        $dialog.Description = '选择漫画导出目录'
        if (Test-Path -LiteralPath $outputBox.Text -PathType Container) { $dialog.SelectedPath = $outputBox.Text }
        if ($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::OK) { $outputBox.Text = $dialog.SelectedPath }
        $dialog.Dispose()
    })

    $export.Add_Click({
        $selectedIndexes = @($list.CheckedIndices | ForEach-Object { [int]$_ })
        if ($selectedIndexes.Count -eq 0) {
            [Windows.Forms.MessageBox]::Show('请至少勾选一部漫画。', '漫画 CBZ / EPUB 导出器', 'OK', 'Information') | Out-Null
            return
        }
        $resolvedOutput = Resolve-OutputPath -LibraryRoot $LibraryRoot -Candidate $outputBox.Text
        $selectedDirectories = @($selectedIndexes | ForEach-Object { $script:CandidateItems[$_] })
        foreach ($directory in $selectedDirectories) {
            if (Test-PathInside -ChildPath $resolvedOutput -ParentPath $directory.FullName) {
                [Windows.Forms.MessageBox]::Show(('输出目录不能放在待导出的漫画内部：' + $directory.Name), '漫画 CBZ / EPUB 导出器', 'OK', 'Error') | Out-Null
                return
            }
        }
        $form.Cursor = [Windows.Forms.Cursors]::WaitCursor
        $export.Enabled = $false
        try {
            $exportMode = if ($epubBook.Checked) { 'Epub' } elseif ($singleBook.Checked) { 'SingleBook' } else { 'PerChapter' }
            $plans = @()
            for ($index = 0; $index -lt $selectedDirectories.Count; $index++) {
                $status.Text = ('正在核验 {0}/{1}：{2}' -f ($index + 1), $selectedDirectories.Count, $selectedDirectories[$index].Name)
                [Windows.Forms.Application]::DoEvents()
                $plans += Get-ComicPlan -ComicDirectory $selectedDirectories[$index]
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
                $results = @(Export-ComicPlan -Plan $plans[$index] -ResolvedOutputPath $resolvedOutput -ExportMode $exportMode -UseCover $coverCheck.Checked)
                $filesCreated += $results.Count
                foreach ($result in $results) { $imagesWritten += $result.ImageCount }
            }
            $status.Text = ('完成：{0} 个导出文件，写入 {1} 个图片条目。' -f $filesCreated, $imagesWritten)
            [Windows.Forms.MessageBox]::Show(($status.Text + "`r`n`r`n输出目录：" + $resolvedOutput), '漫画 CBZ / EPUB 导出器', 'OK', 'Information') | Out-Null
            if ($openCheck.Checked) { Start-Process -FilePath 'explorer.exe' -ArgumentList @($resolvedOutput) }
        }
        catch {
            $status.Text = '导出失败：' + $_.Exception.Message
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, '漫画 CBZ / EPUB 导出器错误', 'OK', 'Error') | Out-Null
        }
        finally {
            $export.Enabled = $true
            $form.Cursor = [Windows.Forms.Cursors]::Default
        }
    })

    & $loadCandidates
    if ($SmokeTest) {
        $timer = New-Object Windows.Forms.Timer
        $timer.Interval = 350
        $timer.Add_Tick({ $timer.Stop(); $form.Close() })
        $form.Add_Shown({ $timer.Start() })
    }
    [void]$form.ShowDialog()
}

try {
    $toolPath = ''
    if (-not [string]::IsNullOrWhiteSpace($env:LOCAL_COMIC_TOOL_PATH)) { $toolPath = $env:LOCAL_COMIC_TOOL_PATH }
    elseif ($null -ne $MyInvocation.MyCommand.PSObject.Properties['Path']) { $toolPath = [string]$MyInvocation.MyCommand.Path }
    $scriptDirectory = if ([string]::IsNullOrWhiteSpace($toolPath)) { (Get-Location).Path } else { [IO.Path]::GetDirectoryName($toolPath) }
    $rootCandidate = if ([string]::IsNullOrWhiteSpace($RootPath)) { $scriptDirectory } elseif ([IO.Path]::IsPathRooted($RootPath)) { $RootPath } else { Join-Path $scriptDirectory $RootPath }
    $resolvedRoot = (Resolve-Path -LiteralPath $rootCandidate).Path
    $resolvedOutput = Resolve-OutputPath -LibraryRoot $resolvedRoot -Candidate $OutputPath

    if ($NonInteractive) {
        if ([string]::IsNullOrWhiteSpace($ComicName)) { throw '非交互模式必须提供 -ComicName。' }
        $comicPath = Join-Path $resolvedRoot $ComicName
        if (-not (Test-Path -LiteralPath $comicPath -PathType Container)) { throw ('漫画文件夹不存在：' + $comicPath) }
        if (Test-PathInside -ChildPath $resolvedOutput -ParentPath $comicPath) { throw '输出目录不能放在待导出的漫画内部。' }
        $plan = Get-ComicPlan -ComicDirectory (Get-Item -LiteralPath $comicPath)
        foreach ($message in $plan.Issues) { Write-Host ('[需确认] ' + $message) -ForegroundColor Yellow }
        foreach ($message in $plan.Warnings) { Write-Host ('[提示] ' + $message) -ForegroundColor DarkYellow }
        Write-Host ('[核验] {0}：{1} 话，{2} 张正文图片，元数据顺序={3}' -f $plan.Name, $plan.ChapterCount, $plan.TotalImages, $plan.MetadataUsed)
        if ($ValidateOnly) { exit 0 }
        if ($plan.Issues.Count -gt 0 -and -not $ForceIssues) { throw '存在需要确认的编号或文件问题；非交互导出请添加 -ForceIssues。' }
        [void](New-Item -ItemType Directory -Path $resolvedOutput -Force)
        $results = @(Export-ComicPlan -Plan $plan -ResolvedOutputPath $resolvedOutput -ExportMode $Mode -UseCover ([bool]$IncludeCover))
        foreach ($result in $results) { Write-Host ('[完成] {0}（{1} 个图片条目）' -f $result.Path, $result.ImageCount) -ForegroundColor Green }
        exit 0
    }

    Show-ExporterWindow -LibraryRoot $resolvedRoot -InitialOutput $resolvedOutput -SmokeTest:$UiSmokeTest
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
        [Windows.Forms.MessageBox]::Show($_.Exception.Message, '漫画 CBZ / EPUB 导出器错误', 'OK', 'Error') | Out-Null
    }
    exit 1
}
