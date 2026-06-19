<#
.SYNOPSIS
    Windows PowerShell 本地文件整理工具。

.DESCRIPTION
    安全原则：
    - 只处理用户明确指定的 Source 文件夹。
    - 默认 dry-run（预演）模式，只生成计划，不移动/复制文件。
    - 不删除文件，不修改文件内容，不覆盖同名文件。
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Source,

    [Parameter(Mandatory = $true)]
    [string]$Target,

    [switch]$Execute,

    [switch]$Copy,

    [string]$Plan = "organize_plan.csv",

    [string]$Log = "organize_log.txt"
)

$ErrorActionPreference = "Stop"

$TypeFolders = [ordered]@{
    "Word文档" = @(".doc", ".docx", ".dot", ".dotx", ".rtf")
    "PDF文件" = @(".pdf")
    "Excel表格" = @(".xls", ".xlsx", ".xlsm", ".xlt", ".xltx", ".csv")
    "PPT演示" = @(".ppt", ".pptx", ".pps", ".ppsx")
    "图片" = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic", ".svg")
    "视频" = @(".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv", ".mpeg", ".mpg", ".webm")
    "音频" = @(".mp3", ".wav", ".flac", ".aac", ".m4a", ".ogg", ".wma")
    "压缩包" = @(".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz")
    "文本文件" = @(".txt", ".md", ".log", ".ini", ".json", ".xml", ".yaml", ".yml")
}
$OtherType = "其他文件"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")

    $line = "{0} - {1} - {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    Write-Host $line
}

function Resolve-Folder {
    param([string]$PathText, [string]$Label)

    $resolved = (Resolve-Path -LiteralPath $PathText).ProviderPath
    if (-not (Test-Path -LiteralPath $resolved -PathType Container)) {
        throw "$Label 不是文件夹：$resolved"
    }
    return $resolved
}

function Get-FileType {
    param([System.IO.FileInfo]$File)

    $extension = $File.Extension.ToLowerInvariant()
    foreach ($folderName in $TypeFolders.Keys) {
        if ($TypeFolders[$folderName] -contains $extension) {
            return $folderName
        }
    }
    return $OtherType
}

function Test-IsInFolder {
    param([string]$Path, [string]$Folder)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullFolder = [System.IO.Path]::GetFullPath($Folder)
    if (-not $fullFolder.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $fullFolder = $fullFolder + [System.IO.Path]::DirectorySeparatorChar
    }
    return $fullPath.StartsWith($fullFolder, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-UniqueTargetPath {
    param([string]$DesiredPath, [hashtable]$ReservedPaths)

    $normalizedDesired = [System.IO.Path]::GetFullPath($DesiredPath).ToLowerInvariant()
    $hasConflict = (Test-Path -LiteralPath $DesiredPath) -or $ReservedPaths.ContainsKey($normalizedDesired)
    if (-not $hasConflict) {
        $ReservedPaths[$normalizedDesired] = $true
        return [pscustomobject]@{ Path = $DesiredPath; HasConflict = $false }
    }

    $parent = [System.IO.Path]::GetDirectoryName($DesiredPath)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DesiredPath)
    $suffix = [System.IO.Path]::GetExtension($DesiredPath)
    $counter = 1

    while ($true) {
        $candidate = Join-Path -Path $parent -ChildPath ("{0}_{1:000}{2}" -f $stem, $counter, $suffix)
        $normalizedCandidate = [System.IO.Path]::GetFullPath($candidate).ToLowerInvariant()
        if ((-not (Test-Path -LiteralPath $candidate)) -and (-not $ReservedPaths.ContainsKey($normalizedCandidate))) {
            $ReservedPaths[$normalizedCandidate] = $true
            return [pscustomobject]@{ Path = $candidate; HasConflict = $true }
        }
        $counter++
    }
}

function Build-Plan {
    param([string]$SourcePath, [string]$TargetPath)

    $reservedPaths = @{}
    $planItems = New-Object System.Collections.Generic.List[object]
    $files = Get-ChildItem -LiteralPath $SourcePath -File -Recurse -Force

    foreach ($file in $files) {
        if (Test-IsInFolder -Path $file.FullName -Folder $TargetPath) {
            Write-Log "跳过目标文件夹内的文件：$($file.FullName)"
            continue
        }

        $fileType = Get-FileType -File $file
        $year = $file.LastWriteTime.Year.ToString()
        $desiredTarget = Join-Path -Path $TargetPath -ChildPath (Join-Path -Path $fileType -ChildPath (Join-Path -Path $year -ChildPath $file.Name))
        $uniqueTarget = Get-UniqueTargetPath -DesiredPath $desiredTarget -ReservedPaths $reservedPaths

        $planItems.Add([pscustomobject]@{
            "原文件路径" = $file.FullName
            "目标文件路径" = $uniqueTarget.Path
            "文件类型" = $fileType
            "文件大小" = $file.Length
            "最后修改时间" = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
            "是否会重名" = if ($uniqueTarget.HasConflict) { "是" } else { "否" }
        }) | Out-Null
    }

    return $planItems
}

try {
    $script:LogPath = [System.IO.Path]::GetFullPath($Log)
    if (Test-Path -LiteralPath $script:LogPath) {
        Remove-Item -LiteralPath $script:LogPath -Force
    }

    $sourcePath = Resolve-Folder -PathText $Source -Label "源文件夹"
    $targetPath = [System.IO.Path]::GetFullPath($Target)
    $planPath = [System.IO.Path]::GetFullPath($Plan)

    Write-Log "源文件夹：$sourcePath"
    Write-Log "目标文件夹：$targetPath"
    Write-Log ("模式：{0}{1}" -f $(if ($Execute) { "正式执行" } else { "预演 dry-run" }), $(if ($Copy) { "（复制）" } else { "（移动）" }))

    $planItems = Build-Plan -SourcePath $sourcePath -TargetPath $targetPath
    if ($planItems.Count -gt 0) {
        $planItems | Export-Csv -LiteralPath $planPath -NoTypeInformation -Encoding UTF8
    } else {
        '"原文件路径","目标文件路径","文件类型","文件大小","最后修改时间","是否会重名"' |
            Set-Content -LiteralPath $planPath -Encoding UTF8
    }
    Write-Log "已生成整理计划：$planPath，共 $($planItems.Count) 个文件"

    if ($Execute) {
        New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
        foreach ($item in $planItems) {
            $destinationFolder = Split-Path -Path $item."目标文件路径" -Parent
            New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
            if ($Copy) {
                Write-Log "复制：$($item.'原文件路径') -> $($item.'目标文件路径')"
                Copy-Item -LiteralPath $item."原文件路径" -Destination $item."目标文件路径" -Force:$false
            } else {
                Write-Log "移动：$($item.'原文件路径') -> $($item.'目标文件路径')"
                Move-Item -LiteralPath $item."原文件路径" -Destination $item."目标文件路径" -Force:$false
            }
        }
        Write-Log "正式执行完成。"
    } else {
        Write-Log "当前是预演 dry-run，未移动、未复制、未删除任何文件。确认计划后添加 -Execute 再正式执行。"
    }

    exit 0
} catch {
    if (-not $script:LogPath) {
        $script:LogPath = [System.IO.Path]::GetFullPath($Log)
    }
    Write-Log "发生错误：$($_.Exception.Message)" "ERROR"
    exit 1
}
