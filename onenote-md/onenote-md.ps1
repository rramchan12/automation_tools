param(
    [Parameter(Position = 0)]
    [ValidateSet('help', 'list', 'export-page', 'export-subtree', 'export-section', 'sync')]
    [string]$Command = 'help',

    [string]$Notebook,
    [string]$Section,
    [string]$Page,
    [Alias('Destination', 'OutputPath')]
    [string]$Out = (Join-Path (Get-Location).Path 'OneNoteMarkdown'),
    [Alias('Log', 'LogPath')]
    [string]$LogFile,
    [switch]$ChangedOnly,
    [switch]$Versioned,
    [switch]$IncludeXml
)

$ErrorActionPreference = 'Stop'

function Write-FatalErrorLog {
    param(
        [string]$Message,
        [string]$ScriptStack
    )

    try {
        if ([string]::IsNullOrWhiteSpace($script:LogFile)) {
            $fallbackLogRoot = Join-Path $script:Out '_logs'
            $script:LogFile = Join-Path $fallbackLogRoot ('onenote-md-fatal-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        }

        $logDirectory = Split-Path -Parent $script:LogFile
        if (-not [string]::IsNullOrWhiteSpace($logDirectory)) {
            New-Item -ItemType Directory -Force -Path $logDirectory | Out-Null
        }

        $lines = New-Object System.Collections.Generic.List[string]
        [void]$lines.Add('OneNote Markdown export validation log')
        [void]$lines.Add(('Generated: {0}' -f (Get-Date).ToString('o')))
        [void]$lines.Add('Status: FAILED')
        [void]$lines.Add(('Command: {0}' -f $script:Command))
        [void]$lines.Add(('Notebook: {0}' -f $(if ([string]::IsNullOrWhiteSpace($script:Notebook)) { '<any>' } else { $script:Notebook })))
        [void]$lines.Add(('Section: {0}' -f $(if ([string]::IsNullOrWhiteSpace($script:Section)) { '<none>' } else { $script:Section })))
        [void]$lines.Add(('Page: {0}' -f $(if ([string]::IsNullOrWhiteSpace($script:Page)) { '<none>' } else { $script:Page })))
        [void]$lines.Add(('Output root argument: {0}' -f $script:Out))
        [void]$lines.Add('Errors: 1')
        [void]$lines.Add('')
        [void]$lines.Add('Fatal error')
        [void]$lines.Add($Message)
        if (-not [string]::IsNullOrWhiteSpace($ScriptStack)) {
            [void]$lines.Add('')
            [void]$lines.Add('Script stack')
            [void]$lines.Add($ScriptStack)
        }

        $lines | Set-Content -LiteralPath $script:LogFile -Encoding UTF8
        Write-Output ('Log: {0}' -f $script:LogFile)
    }
    catch {
        Write-Output ('Failed to write validation log: {0}' -f $_.Exception.Message)
    }
}

trap {
    $fatalMessage = "Line {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message
    Write-FatalErrorLog $fatalMessage $_.ScriptStackTrace
    Write-Error $fatalMessage
    if ($_.ScriptStackTrace) {
        Write-Error $_.ScriptStackTrace
    }
    exit 1
}

function Show-Help {
    @'
onenote-md.ps1 - OneNote to readable Markdown exporter

Commands:
  help
  list
  export-page    -Section "<SectionName>" -Page "<PageName>"
  export-subtree -Section "<SectionName>" -Page "<PageName>"
  export-section -Section "<SectionName>"
  sync           -Section "<SectionName>" -ChangedOnly

Common parameters:
  -Notebook      Optional notebook name filter.
  -Section       OneNote section name.
  -Page          Page name for page/subtree export.
  -Out           Output root folder. Aliases: -Destination, -OutputPath.
  -LogFile       Optional validation log path. Aliases: -Log, -LogPath.
  -ChangedOnly   Skip pages whose OneNote lastModifiedTime has not changed.
  -Versioned     Also write into a timestamped versions folder.
  -IncludeXml    Save raw OneNote page XML next to Markdown for debugging.

Output:
  <Out>\<Section>\latest\*.md
  <Out>\<Section>\versions\<yyyyMMdd-HHmmss>\*.md when -Versioned is set
  <Out>\<Section>\.onenote-md-manifest.json
  <Out>\<Section>\logs\onenote-md-<yyyyMMdd-HHmmss>.log unless -LogFile is set
'@
}

function New-OneNoteApplication {
    try {
        return New-Object -ComObject OneNote.Application
    }
    catch {
        throw "Unable to start OneNote COM automation. Open OneNote desktop once and make sure the target notebook is available. $($_.Exception.Message)"
    }
}

function Get-OneNoteHierarchy {
    param($OneNote)

    $hierarchyXml = ''
    $oneNote.GetHierarchy('', 4, [ref]$hierarchyXml)
    [xml]$hierarchy = $hierarchyXml
    return $hierarchy
}

function New-NamespaceManager {
    param([xml]$Xml)

    $ns = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
    $ns.AddNamespace('one', 'http://schemas.microsoft.com/office/onenote/2013/onenote')
    return $ns
}

function Get-SafeFileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = 'Untitled'
    }

    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $Name = $Name.Replace($char, '-')
    }

    $Name = ($Name -replace '\s+', ' ').Trim()
    if ($Name.Length -gt 100) {
        $Name = $Name.Substring(0, 100).Trim()
    }
    return $Name
}

function Convert-HtmlFragmentToMarkdownText {
    param([string]$Html)

    if ($null -eq $Html) {
        return ''
    }

    $text = [System.Net.WebUtility]::HtmlDecode($Html)
    $text = $text -replace "`r`n", "`n" -replace "`r", "`n"
    $text = [regex]::Replace($text, '<br\s*/?>', "`n", 'IgnoreCase')
    $text = [regex]::Replace($text, '<a\s+[^>]*href\s*=\s*[''"]([^''"]+)[''"][^>]*>(.*?)</a>', {
        param($match)
        $url = $match.Groups[1].Value
        $label = [regex]::Replace($match.Groups[2].Value, '<[^>]+>', '')
        $label = [System.Net.WebUtility]::HtmlDecode($label).Trim()
        if ([string]::IsNullOrWhiteSpace($label)) {
            $label = $url
        }
        return '[' + $label + '](' + $url + ')'
    }, 'IgnoreCase,Singleline')

    $text = [regex]::Replace($text, '<span\s+[^>]*style\s*=\s*[''"]([^''"]*)[''"][^>]*>(.*?)</span>', {
        param($match)
        $style = $match.Groups[1].Value.ToLowerInvariant()
        $label = [regex]::Replace($match.Groups[2].Value, '<[^>]+>', '')
        $label = [System.Net.WebUtility]::HtmlDecode($label).Trim()
        if ($style.Contains('font-weight:bold')) {
            return '**' + $label + '**'
        }
        return $label
    }, 'IgnoreCase,Singleline')

    $text = [regex]::Replace($text, '<[^>]+>', '')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    $text = $text -replace '[ \t]+', ' '
    return $text.Trim()
}

function Get-MarkdownLink {
    param([string]$Line)

    $lineValue = $Line.Trim()
    if ($lineValue.StartsWith('[') -and $lineValue.EndsWith(']') -and $lineValue -match '^\[(\[[^\]]+\]\(.+\))\]$') {
        $lineValue = $Matches[1]
    }

    $match = [regex]::Match($lineValue, '^\[([^\]]+)\]\((.+)\)$')
    if (-not $match.Success) {
        return $null
    }

    return [pscustomobject]@{
        Text = $match.Groups[1].Value
        Url  = $match.Groups[2].Value
    }
}

function Find-Section {
    param(
        [xml]$Hierarchy,
        [string]$NotebookName,
        [string]$SectionName
    )

    $sectionNodes = $Hierarchy.SelectNodes('//*[local-name()="Section"]')
    $sections = New-Object System.Collections.Generic.List[object]
    for ($sectionIndex = 0; $sectionIndex -lt $sectionNodes.Count; $sectionIndex++) {
        [void]$sections.Add($sectionNodes.Item($sectionIndex))
    }
    $matches = New-Object System.Collections.Generic.List[object]

    foreach ($sectionNode in $sections) {
        $sectionMatches = [string]::IsNullOrWhiteSpace($SectionName) -or $sectionNode.GetAttribute('name') -eq $SectionName
        if (-not $sectionMatches) {
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($NotebookName)) {
            $notebookNode = $sectionNode.ParentNode
            while ($null -ne $notebookNode -and $notebookNode.LocalName -ne 'Notebook') {
                $notebookNode = $notebookNode.ParentNode
            }
            if ($null -eq $notebookNode -or $notebookNode.GetAttribute('name') -ne $NotebookName) {
                continue
            }
        }

        [void]$matches.Add($sectionNode)
    }

    if ($matches.Count -eq 0) {
        throw "Section not found: $SectionName"
    }
    if ($matches.Count -gt 1) {
        $names = ($matches | ForEach-Object { $_.GetAttribute('path') }) -join "`n"
        throw "Multiple sections matched. Specify -Notebook as well.`n$names"
    }

    return [pscustomobject]@{ Node = $matches[0] }
}

function Get-SectionPages {
    param($SectionNode)

    $pageNodes = $SectionNode.SelectNodes('*[local-name()="Page"]')
    $pages = New-Object System.Collections.Generic.List[object]
    for ($pageIndex = 0; $pageIndex -lt $pageNodes.Count; $pageIndex++) {
        [void]$pages.Add($pageNodes.Item($pageIndex))
    }
    return ,$pages.ToArray()
}

function Get-PageSubtree {
    param(
        [object[]]$Pages,
        [string]$PageName
    )

    $rootIndex = -1
    for ($i = 0; $i -lt $Pages.Count; $i++) {
        if ($Pages[$i].GetAttribute('name') -eq $PageName) {
            $rootIndex = $i
            break
        }
    }

    if ($rootIndex -lt 0) {
        throw "Page not found: $PageName"
    }

    $rootLevel = [int]$Pages[$rootIndex].GetAttribute('pageLevel')
    $subtree = New-Object System.Collections.Generic.List[object]
    [void]$subtree.Add($Pages[$rootIndex])

    for ($i = $rootIndex + 1; $i -lt $Pages.Count; $i++) {
        $level = [int]$Pages[$i].GetAttribute('pageLevel')
        if ($level -le $rootLevel) {
            break
        }
        [void]$subtree.Add($Pages[$i])
    }

    return $subtree.ToArray()
}

function Read-Manifest {
    param([string]$ManifestPath)

    if (-not (Test-Path -LiteralPath $ManifestPath)) {
        return @{}
    }

    $json = Get-Content -LiteralPath $ManifestPath -Raw
    if ([string]::IsNullOrWhiteSpace($json)) {
        return @{}
    }

    $items = ConvertFrom-Json $json
    $manifest = @{}
    foreach ($item in $items.pages) {
        $manifest[$item.id] = $item
    }
    return $manifest
}

function Write-Manifest {
    param(
        [string]$ManifestPath,
        [object[]]$Records
    )

    $payload = [pscustomobject]@{
        generatedAt = (Get-Date).ToString('o')
        pages       = @($Records)
    }
    $payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8
}

function Export-OneNotePage {
    param(
        $OneNote,
        $PageNode,
        [string]$FilePath,
        [string]$SectionName,
        [string]$ParentName,
        [switch]$IncludeRawXml
    )

    $pageId = $PageNode.GetAttribute('ID')
    $pageName = $PageNode.GetAttribute('name')
    $lastModified = $PageNode.GetAttribute('lastModifiedTime')
    $pageLevel = $PageNode.GetAttribute('pageLevel')

    $pageXml = ''
    $OneNote.GetPageContent($pageId, [ref]$pageXml, 0)
    [xml]$pageDocument = $pageXml

    if ($IncludeRawXml) {
        $xmlPath = [System.IO.Path]::ChangeExtension($FilePath, '.xml')
        $pageXml | Set-Content -LiteralPath $xmlPath -Encoding UTF8
    }

    $titleNode = $pageDocument.SelectSingleNode('//*[local-name()="Title"]//*[local-name()="T"]')
    if ($null -ne $titleNode) {
        $title = Convert-HtmlFragmentToMarkdownText $titleNode.InnerText
    }
    else {
        $title = $pageName
    }
    if ([string]::IsNullOrWhiteSpace($title)) {
        $title = $pageName
    }

    $rawItems = New-Object System.Collections.Generic.List[string]
    $contentNodes = $pageDocument.SelectNodes('//*[local-name()="Outline"]//*[local-name()="OE"]/*[local-name()="T"]')
    foreach ($contentNode in $contentNodes) {
        $converted = Convert-HtmlFragmentToMarkdownText $contentNode.InnerText
        foreach ($part in ($converted -split "`n")) {
            $part = $part.Trim()
            if (-not [string]::IsNullOrWhiteSpace($part)) {
                [void]$rawItems.Add($part)
            }
        }
    }

    $linkRefs = New-Object System.Collections.Generic.List[object]
    function Add-LinkReference {
        param([string]$Text, [string]$Url)

        [void]$linkRefs.Add([pscustomobject]@{ Text = $Text; Url = $Url })
        return ('[{0}][{1}]' -f $Text, $linkRefs.Count)
    }

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('# ' + $title)
    [void]$lines.Add('')
    [void]$lines.Add(('> Source: OneNote / {0} / {1}' -f $SectionName, $pageName))
    if (-not [string]::IsNullOrWhiteSpace($ParentName)) {
        [void]$lines.Add(('> Parent: {0}' -f $ParentName))
    }
    [void]$lines.Add(('> Last modified: {0}' -f $lastModified))
    [void]$lines.Add('')

    $i = 0
    while ($i -lt $rawItems.Count) {
        $item = $rawItems[$i].Trim()
        if ($item -match '^\*\*(.+)\*\*$') {
            if ($lines[$lines.Count - 1] -ne '') {
                [void]$lines.Add('')
            }
            [void]$lines.Add('## ' + $Matches[1].Trim())
            [void]$lines.Add('')
            $i++
            continue
        }

        $currentLink = Get-MarkdownLink $item
        if (($i + 1) -lt $rawItems.Count) {
            $nextLink = Get-MarkdownLink $rawItems[$i + 1]
        }
        else {
            $nextLink = $null
        }

        if ($null -eq $currentLink -and $null -ne $nextLink) {
            [void]$lines.Add('- ' + $item + ': ' + (Add-LinkReference $nextLink.Text $nextLink.Url))
            $i += 2
            continue
        }

        if ($null -ne $currentLink) {
            [void]$lines.Add('- ' + (Add-LinkReference $currentLink.Text $currentLink.Url))
        }
        else {
            [void]$lines.Add($item)
        }
        $i++
    }

    if ($linkRefs.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add('## Links')
        [void]$lines.Add('')
        for ($j = 0; $j -lt $linkRefs.Count; $j++) {
            [void]$lines.Add(('[{0}]: {1}' -f ($j + 1), $linkRefs[$j].Url))
        }
    }

    $folder = Split-Path -Parent $FilePath
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
    $lines | Set-Content -LiteralPath $FilePath -Encoding UTF8

    return [pscustomobject]@{
        id           = $pageId
        name         = $pageName
        level        = [int]$pageLevel
        lastModified = $lastModified
        file         = [System.IO.Path]::GetFileName($FilePath)
        lines        = $lines.Count
    }
}

function Export-Pages {
    param(
        $OneNote,
        [object[]]$Pages,
        [string]$SectionName,
        [string]$TargetRoot,
        [hashtable]$ExistingManifest,
        [switch]$SkipUnchanged,
        [switch]$IncludeRawXml
    )

    New-Item -ItemType Directory -Force -Path $TargetRoot | Out-Null

    $results = New-Object System.Collections.Generic.List[object]
    $records = New-Object System.Collections.Generic.List[object]
    $rootLevel = [int]$Pages[0].GetAttribute('pageLevel')
    $parentStack = @{}

    for ($i = 0; $i -lt $Pages.Count; $i++) {
        $pageNode = $Pages[$i]
        $pageId = $pageNode.GetAttribute('ID')
        $pageName = $pageNode.GetAttribute('name')
        $pageLevel = [int]$pageNode.GetAttribute('pageLevel')
        $lastModified = $pageNode.GetAttribute('lastModifiedTime')

        $parentName = ''
        for ($level = $pageLevel - 1; $level -ge 1; $level--) {
            if ($parentStack.ContainsKey($level)) {
                $parentName = $parentStack[$level]
                break
            }
        }
        $parentStack[$pageLevel] = $pageName

        if ($i -eq 0) {
            $prefix = '01'
        }
        elseif ($pageLevel -gt $rootLevel) {
            $prefix = '01.' + ('{0:D2}' -f $i)
        }
        else {
            $prefix = '{0:D2}' -f ($i + 1)
        }

        $fileName = ('{0}-{1}.md' -f $prefix, (Get-SafeFileName $pageName))
        $filePath = Join-Path $TargetRoot $fileName

        $unchanged = $false
        if ($SkipUnchanged -and $ExistingManifest.ContainsKey($pageId)) {
            $unchanged = $ExistingManifest[$pageId].lastModified -eq $lastModified
        }

        if ($unchanged -and (Test-Path -LiteralPath $filePath)) {
            $record = [pscustomobject]@{
                id           = $pageId
                name         = $pageName
                level        = $pageLevel
                lastModified = $lastModified
                file         = $fileName
                lines        = $ExistingManifest[$pageId].lines
                status       = 'unchanged'
                error        = ''
            }
        }
        else {
            try {
                $record = Export-OneNotePage $OneNote $pageNode $filePath $SectionName $parentName -IncludeRawXml:$IncludeRawXml
                $record | Add-Member -NotePropertyName status -NotePropertyValue 'exported'
                $record | Add-Member -NotePropertyName error -NotePropertyValue ''
            }
            catch {
                $errorFile = [System.IO.Path]::ChangeExtension($filePath, '.error.txt')
                $_.Exception.Message | Set-Content -LiteralPath $errorFile -Encoding UTF8
                $record = [pscustomobject]@{
                    id           = $pageId
                    name         = $pageName
                    level        = $pageLevel
                    lastModified = $lastModified
                    file         = $fileName
                    lines        = 0
                    status       = 'error'
                    error        = $_.Exception.Message
                    errorFile    = [System.IO.Path]::GetFileName($errorFile)
                }
            }
        }

        [void]$records.Add($record)
        [void]$results.Add($record)
    }

    $index = New-Object System.Collections.Generic.List[string]
    [void]$index.Add('# ' + $SectionName)
    [void]$index.Add('')
    [void]$index.Add('Generated from OneNote.')
    [void]$index.Add('')
    [void]$index.Add('## Pages')
    [void]$index.Add('')

    foreach ($record in $records) {
        if ($record.level -gt $rootLevel) {
            $indent = '  - '
        }
        else {
            $indent = '- '
        }
        [void]$index.Add(('{0}[{1}]({2})' -f $indent, $record.name, $record.file))
    }

    [void]$index.Add('')
    [void]$index.Add('## Export details')
    [void]$index.Add('')
    [void]$index.Add('| Page | Level | Status | Lines | Last modified | File | Error |')
    [void]$index.Add('|---|---:|---|---:|---|---|---|')
    foreach ($record in $records) {
        $escapedName = $record.name -replace '\|', '\|'
        $escapedError = ($record.error -replace '\|', '\|' -replace "`r?`n", ' ')
        [void]$index.Add(('| {0} | {1} | {2} | {3} | {4} | `{5}` | {6} |' -f $escapedName, $record.level, $record.status, $record.lines, $record.lastModified, $record.file, $escapedError))
    }

    $index | Set-Content -LiteralPath (Join-Path $TargetRoot 'index.md') -Encoding UTF8
    return $results.ToArray()
}

function Get-OutputRoots {
    param(
        [string]$OutputRoot,
        [string]$SectionName,
        [switch]$UseVersioned
    )

    $sectionRoot = Join-Path $OutputRoot (Get-SafeFileName $SectionName)
    $latestRoot = Join-Path $sectionRoot 'latest'
    $roots = New-Object System.Collections.Generic.List[string]
    [void]$roots.Add($latestRoot)

    if ($UseVersioned) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        [void]$roots.Add((Join-Path (Join-Path $sectionRoot 'versions') $stamp))
    }

    return [pscustomobject]@{
        SectionRoot = $sectionRoot
        LatestRoot  = $latestRoot
        Roots       = @($roots)
        Manifest    = Join-Path $sectionRoot '.onenote-md-manifest.json'
    }
}

function Write-RunLog {
    param(
        [string]$Path,
        [string]$CommandName,
        [string]$NotebookName,
        [string]$SectionName,
        [string]$PageName,
        [string]$OutputRoot,
        [object]$OutputRoots,
        [object[]]$LatestResults,
        [int]$PagesConsidered,
        [int]$Exported,
        [int]$Unchanged,
        [int]$Errors,
        [switch]$UsedChangedOnly,
        [switch]$UsedVersioned
    )

    $logDirectory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Force -Path $logDirectory | Out-Null

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('OneNote Markdown export validation log')
    [void]$lines.Add(('Generated: {0}' -f (Get-Date).ToString('o')))
    [void]$lines.Add(('Status: {0}' -f $(if ($Errors -gt 0) { 'COMPLETED_WITH_ERRORS' } else { 'SUCCESS' })))
    [void]$lines.Add(('Command: {0}' -f $CommandName))
    [void]$lines.Add(('Notebook: {0}' -f $(if ([string]::IsNullOrWhiteSpace($NotebookName)) { '<any>' } else { $NotebookName })))
    [void]$lines.Add(('Section: {0}' -f $SectionName))
    [void]$lines.Add(('Page: {0}' -f $(if ([string]::IsNullOrWhiteSpace($PageName)) { '<none>' } else { $PageName })))
    [void]$lines.Add(('Output root argument: {0}' -f $OutputRoot))
    [void]$lines.Add(('Latest folder: {0}' -f $OutputRoots.LatestRoot))
    [void]$lines.Add(('Manifest: {0}' -f $OutputRoots.Manifest))
    [void]$lines.Add(('Changed only: {0}' -f [bool]$UsedChangedOnly))
    [void]$lines.Add(('Versioned: {0}' -f [bool]$UsedVersioned))
    [void]$lines.Add(('Pages considered: {0}' -f $PagesConsidered))
    [void]$lines.Add(('Exported: {0}' -f $Exported))
    [void]$lines.Add(('Unchanged: {0}' -f $Unchanged))
    [void]$lines.Add(('Errors: {0}' -f $Errors))
    [void]$lines.Add('')
    [void]$lines.Add('Validation')
    [void]$lines.Add(('Latest index exists: {0}' -f (Test-Path -LiteralPath (Join-Path $OutputRoots.LatestRoot 'index.md'))))
    [void]$lines.Add(('Manifest exists: {0}' -f (Test-Path -LiteralPath $OutputRoots.Manifest)))
    [void]$lines.Add('')
    [void]$lines.Add('Pages')
    [void]$lines.Add('Status | Level | Lines | File exists | Last modified | Page | File | Error')
    [void]$lines.Add('---|---:|---:|---|---|---|---|---')

    foreach ($result in $LatestResults) {
        $filePath = Join-Path $OutputRoots.LatestRoot $result.file
        $errorMessage = $result.error -replace '\|', '\|' -replace "`r?`n", ' '
        [void]$lines.Add(('{0} | {1} | {2} | {3} | {4} | {5} | {6} | {7}' -f $result.status, $result.level, $result.lines, (Test-Path -LiteralPath $filePath), $result.lastModified, $result.name, $filePath, $errorMessage))
    }

    $lines | Set-Content -LiteralPath $Path -Encoding UTF8
}

if ($Command -eq 'help') {
    Show-Help
    exit 0
}

$oneNote = New-OneNoteApplication
$hierarchyXml = ''
$oneNote.GetHierarchy('', 4, [ref]$hierarchyXml)
[xml]$hierarchy = $hierarchyXml
if ($Command -eq 'list') {
    $sectionNodes = $hierarchy.SelectNodes('//*[local-name()="Section"]')
    $rows = New-Object System.Collections.Generic.List[string]
    for ($sectionIndex = 0; $sectionIndex -lt $sectionNodes.Count; $sectionIndex++) {
        $sectionNode = $sectionNodes.Item($sectionIndex)
        $notebookNode = $sectionNode.ParentNode
        while ($null -ne $notebookNode -and $notebookNode.LocalName -ne 'Notebook') {
            $notebookNode = $notebookNode.ParentNode
        }

        if ($null -ne $notebookNode) {
            $notebookName = $notebookNode.GetAttribute('name')
        }
        else {
            $notebookName = ''
        }

        if (-not [string]::IsNullOrWhiteSpace($Notebook) -and $notebookName -ne $Notebook) {
            continue
        }

        $pageCount = $sectionNode.SelectNodes('*[local-name()="Page"]').Count
        $sectionNameForList = $sectionNode.GetAttribute('name')
        $sectionPathForList = $sectionNode.GetAttribute('path')
        [void]$rows.Add(('{0} | {1} | {2} | {3}' -f $notebookName, $sectionNameForList, $pageCount, $sectionPathForList))
    }

    if ($rows.Count -eq 0) {
        Write-Output 'No OneNote sections found. Open OneNote desktop and make sure notebooks are synced, then retry.'
    }
    else {
        Write-Output 'Notebook | Section | Pages | Path'
        Write-Output '---|---|---:|---'
        foreach ($row in $rows) {
            Write-Output $row
        }
    }
    exit 0
}

if ([string]::IsNullOrWhiteSpace($Section)) {
    throw '-Section is required.'
}

$sectionNodesForLookup = $hierarchy.SelectNodes('//*[local-name()="Section"]')
$sectionMatchesForLookup = New-Object System.Collections.Generic.List[int]
for ($lookupIndex = 0; $lookupIndex -lt $sectionNodesForLookup.Count; $lookupIndex++) {
    $candidateSection = $sectionNodesForLookup.Item($lookupIndex)
    if ($candidateSection.GetAttribute('name') -ne $Section) {
        continue
    }

    if (-not [string]::IsNullOrWhiteSpace($Notebook)) {
        $candidateNotebook = $candidateSection.ParentNode
        while ($null -ne $candidateNotebook -and $candidateNotebook.LocalName -ne 'Notebook') {
            $candidateNotebook = $candidateNotebook.ParentNode
        }
        if ($null -eq $candidateNotebook -or $candidateNotebook.GetAttribute('name') -ne $Notebook) {
            continue
        }
    }

    [void]$sectionMatchesForLookup.Add($lookupIndex)
}

if ($sectionMatchesForLookup.Count -eq 0) {
    throw "Section not found: $Section"
}
if ($sectionMatchesForLookup.Count -gt 1) {
    $paths = ($sectionMatchesForLookup | ForEach-Object { $sectionNodesForLookup.Item($_).GetAttribute('path') }) -join "`n"
    throw "Multiple sections matched. Specify -Notebook as well.`n$paths"
}

$selectedSectionNode = $sectionNodesForLookup.Item($sectionMatchesForLookup[0])
if (-not ($selectedSectionNode -is [System.Xml.XmlElement])) {
    $actualType = if ($null -eq $selectedSectionNode) { '<null>' } else { $selectedSectionNode.GetType().FullName }
    throw "Internal error: section lookup returned $actualType instead of System.Xml.XmlElement. Value: $selectedSectionNode"
}
$sectionName = $selectedSectionNode.GetAttribute('name')
$pages = Get-SectionPages $selectedSectionNode

if ($Command -eq 'export-page' -or $Command -eq 'export-subtree') {
    if ([string]::IsNullOrWhiteSpace($Page)) {
        throw '-Page is required.'
    }

    if ($Command -eq 'export-page') {
        $matched = @($pages | Where-Object { $_.GetAttribute('name') -eq $Page })
        if ($matched.Count -eq 0) {
            throw "Page not found: $Page"
        }
        if ($matched.Count -gt 1) {
            throw "Multiple pages matched: $Page"
        }
        $pagesToExport = @($matched[0])
    }
    else {
        $pagesToExport = Get-PageSubtree $pages $Page
    }
}
elseif ($Command -eq 'export-section' -or $Command -eq 'sync') {
    $pagesToExport = $pages
    if ($Command -eq 'sync') {
        $ChangedOnly = $true
    }
}
else {
    throw "Unsupported command: $Command"
}

$outputRoots = Get-OutputRoots $Out $sectionName -UseVersioned:$Versioned
New-Item -ItemType Directory -Force -Path $outputRoots.SectionRoot | Out-Null
$manifest = Read-Manifest $outputRoots.Manifest

$allResults = @()
$latestResults = @()
foreach ($root in $outputRoots.Roots) {
    $results = Export-Pages $oneNote $pagesToExport $sectionName $root $manifest -SkipUnchanged:$ChangedOnly -IncludeRawXml:$IncludeXml
    if ($root -eq $outputRoots.LatestRoot) {
        $latestResults = @($results)
    }
    $allResults += $results
}

Write-Manifest $outputRoots.Manifest $latestResults

$exported = @($latestResults | Where-Object { $_.status -eq 'exported' }).Count
$unchanged = @($latestResults | Where-Object { $_.status -eq 'unchanged' }).Count
$errors = @($latestResults | Where-Object { $_.status -eq 'error' }).Count

if ([string]::IsNullOrWhiteSpace($LogFile)) {
    $logRoot = Join-Path $outputRoots.SectionRoot 'logs'
    $LogFile = Join-Path $logRoot ('onenote-md-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
}

Write-RunLog `
    -Path $LogFile `
    -CommandName $Command `
    -NotebookName $Notebook `
    -SectionName $sectionName `
    -PageName $Page `
    -OutputRoot $Out `
    -OutputRoots $outputRoots `
    -LatestResults $latestResults `
    -PagesConsidered $pagesToExport.Count `
    -Exported $exported `
    -Unchanged $unchanged `
    -Errors $errors `
    -UsedChangedOnly:$ChangedOnly `
    -UsedVersioned:$Versioned

Write-Output ('Section: {0}' -f $sectionName)
Write-Output ('Pages considered: {0}' -f $pagesToExport.Count)
Write-Output ('Exported: {0}' -f $exported)
Write-Output ('Unchanged: {0}' -f $unchanged)
Write-Output ('Errors: {0}' -f $errors)
Write-Output ('Latest: {0}' -f $outputRoots.LatestRoot)
Write-Output ('Manifest: {0}' -f $outputRoots.Manifest)
Write-Output ('Log: {0}' -f $LogFile)

if ($errors -gt 0) {
    exit 2
}
