<# 
.SYNOPSIS
  Compare text files in two folder trees.

.DESCRIPTION
  Scans two roots (Left/Right), matches files by relative path and extension,
  and reports: OnlyInLeft, OnlyInRight, Same, Different.
  For Different files, can emit per-file split outputs (LEFT/RIGHT) and (optionally)
  a combined unified-style diff. Optionally auto-opens a GUI diff viewer.

.PARAMETER LeftRoot
  Left/top folder path.

.PARAMETER RightRoot
  Right/top folder path.

.PARAMETER Extensions
  One or more file extensions to treat as "text". Defaults include 
  .txt,.csv,.log,.json,.xml,.ps1,.psm1,.psd1,.ini,.cfg,.conf,.sql

.PARAMETER ExcludePattern
  Wildcard pattern(s) (array) to exclude relative paths, e.g. '*\bin\*','*.tmp'

.PARAMETER Encoding
  Text encoding hint for reading files. Default: utf8. (If reading fails, falls back to default).

.PARAMETER WriteDiffs
  If set, writes outputs for changed files under a _diffs folder (or in -DiffOutputDir).

.PARAMETER SplitDiffsOnly
  With -WriteDiffs, writes only split files: *.left.txt and *.right.txt (no combined .diff.txt).

.PARAMETER DiffOutputDir
  Optional path where per-file outputs are written.

.PARAMETER ReportCsv
  Optional path to write a CSV summary of results.

.PARAMETER Fast
  If set, uses file size + SHA256 hashes for change detection (skips line-by-line compare
  when obviously identical; still produces split outputs when -WriteDiffs is set).

.PARAMETER OpenDiffViewer
  After writing split outputs for changed files, automatically open a diff viewer.

.PARAMETER DiffViewer
  Which viewer to use when -OpenDiffViewer is set. Options: Auto, VisualStudio, VSCode.
  Default: Auto (prefers Visual Studio if found, else VS Code).

.PARAMETER MaxOpen
  Max number of changed files to open in the diff viewer (prevents opening too many windows).
  Default: 5
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][ValidateScript({ Test-Path $_ -PathType Container })]
  [string]$LeftRoot,

  [Parameter(Mandatory=$true)][ValidateScript({ Test-Path $_ -PathType Container })]
  [string]$RightRoot,

  [string[]]$Extensions = @('.txt','.csv','.log','.json','.xml','.ps1','.psm1','.psd1','.ini','.cfg','.conf','.sql'),

  [string[]]$ExcludePattern = @(),

  [ValidateSet('utf8','ascii','unicode','utf7','utf32','bigendianunicode','oem')]
  [string]$Encoding = 'utf8',

  [switch]$WriteDiffs,
  [switch]$SplitDiffsOnly,
  [string]$DiffOutputDir,
  [string]$ReportCsv,

  [switch]$Fast,

  [switch]$OpenDiffViewer,
  [ValidateSet('Auto','VisualStudio','VSCode')]
  [string]$DiffViewer = 'Auto',
  [int]$MaxOpen = 5
)

function Resolve-RelativePath {
  param([string]$Root,[string]$FullPath)
  $rootFull = [System.IO.Path]::GetFullPath($Root)
  $fileFull = [System.IO.Path]::GetFullPath($FullPath)
  if ($fileFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
    $rel = $fileFull.Substring($rootFull.Length).TrimStart('\','/')
    return $rel
  }
  return $FullPath
}

function Get-TextFiles {
  param([string]$Root,[string[]]$Exts,[string[]]$Exclude)

  # Build a case-insensitive set of extensions, normalizing to start with a dot
  $extSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($e in $Exts) {
    $ext = $e
    if ($null -ne $ext -and $ext.Length -gt 0) {
      if (-not $ext.StartsWith('.')) { $ext = ".$ext" }
      [void]$extSet.Add($ext)
    }
  }

  $files = Get-ChildItem -Path $Root -Recurse -File -ErrorAction Stop |
    Where-Object { $extSet.Contains([System.IO.Path]::GetExtension($_.FullName)) } |
    ForEach-Object {
      $rel = Resolve-RelativePath -Root $Root -FullPath $_.FullName

      # Exclusions are matched against the relative path
      $excluded = $false
      foreach ($p in $Exclude) { if ($rel -like $p) { $excluded = $true; break } }

      if (-not $excluded) {
        [PSCustomObject]@{
          RelativePath = $rel
          FullPath     = $_.FullName
          Length       = $_.Length
          LastWrite    = $_.LastWriteTimeUtc
        }
      }
    }
  return $files
}

function Get-FileHashSafe {
  param([string]$Path)
  try { (Get-FileHash -Algorithm SHA256 -LiteralPath $Path -ErrorAction Stop).Hash }
  catch { $null }
}

function Read-LinesSafe {
  param([string]$Path,[string]$Enc)
  try {
    Get-Content -LiteralPath $Path -Encoding $Enc -ErrorAction Stop
  } catch {
    try { Get-Content -LiteralPath $Path -ErrorAction Stop } catch { @() }
  }
}

function Write-UnifiedDiff {
  [CmdletBinding()]
  param(
    [string[]]$LeftLines,
    [string[]]$RightLines,
    [string]$LeftLabel,
    [string]$RightLabel,
    [string]$OutFile
  )
  # Simple unified diff (context-less): '-', '+', ' ' lines
  $max = [Math]::Max($LeftLines.Count, $RightLines.Count)
  $out = New-Object System.Collections.Generic.List[string]
  $out.Add("--- $LeftLabel")
  $out.Add("+++ $RightLabel")
  for ($i=0; $i -lt $max; $i++) {
    $l = if ($i -lt $LeftLines.Count) { $LeftLines[$i] } else { $null }
    $r = if ($i -lt $RightLines.Count) { $RightLines[$i] } else { $null }
    if ($l -eq $r) {
      $out.Add(" $l")
    } elseif ($l -ne $null -and $r -ne $null) {
      $out.Add("-$l")
      $out.Add("+$r")
    } elseif ($l -ne $null) {
      $out.Add("-$l")
    } else {
      $out.Add("+$r")
    }
  }
  $out | Set-Content -LiteralPath $OutFile -Encoding UTF8
}

# ---- Diff viewer helpers ----

function Test-Executable {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
  try { return (Test-Path -LiteralPath $Path -PathType Leaf) } catch { return $false }
}

function Get-CodePath {
  # Try PATH first
  $pathCmd = (Get-Command code -ErrorAction SilentlyContinue)
  if ($pathCmd) { return $pathCmd.Source }

  # Common install locations (user installer and system installer)
  $cand = @(
    Join-Path $env:LOCALAPPDATA "Programs\Microsoft VS Code\Code.exe"),
    "C:\Program Files\Microsoft VS Code\Code.exe",
    "C:\Program Files (x86)\Microsoft VS Code\Code.exe"
  foreach ($c in $cand) { if (Test-Executable $c) { return $c } }
  return $null
}

function Get-DevenvPath {
  # Prefer vswhere if present
  $vswhere = "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
  if (Test-Executable $vswhere) {
    try {
      $line = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property productPath 2>$null
      if ($line -and (Test-Executable $line)) { return $line }
    } catch { }
  }
  # Fallback to common 2022/2019 editions
  $roots = @("C:\Program Files\Microsoft Visual Studio\2022",
             "C:\Program Files (x86)\Microsoft Visual Studio\2019")
  $editions = @("Community","Professional","Enterprise")
  foreach ($r in $roots) {
    foreach ($e in $editions) {
      $p = Join-Path (Join-Path $r $e) "Common7\IDE\devenv.exe"
      if (Test-Executable $p) { return $p }
    }
  }
  return $null
}

function Resolve-DiffViewer {
  param([string]$Preference) # Auto | VisualStudio | VSCode
  $pref = $Preference
  if ([string]::IsNullOrWhiteSpace($pref)) { $pref = 'Auto' }

  if ($pref -eq 'VisualStudio') {
    $vs = Get-DevenvPath
    if ($vs) { return @{ Name='VisualStudio'; Path=$vs } }
    return $null
  }

  if ($pref -eq 'VSCode') {
    $code = Get-CodePath
    if ($code) { return @{ Name='VSCode'; Path=$code } }
    return $null
  }

  # Auto: prefer Visual Studio, then VS Code
  $vsAuto = Get-DevenvPath
  if ($vsAuto) { return @{ Name='VisualStudio'; Path=$vsAuto } }

  $codeAuto = Get-CodePath
  if ($codeAuto) { return @{ Name='VSCode'; Path=$codeAuto } }

  return $null
}

function Open-Diff {
  param(
    [string]$ViewerName,  # VisualStudio | VSCode
    [string]$ViewerPath,
    [string]$LeftFile,
    [string]$RightFile
  )
  if ($ViewerName -eq 'VisualStudio') {
    # devenv.exe /diff "left" "right"
    Start-Process -FilePath $ViewerPath -ArgumentList @('/diff', $LeftFile, $RightFile) | Out-Null
    return
  }
  if ($ViewerName -eq 'VSCode') {
    # code --diff "left" "right"
    Start-Process -FilePath $ViewerPath -ArgumentList @('--diff', $LeftFile, $RightFile) | Out-Null
    return
  }
}

# ---- Prep ----
$LeftRoot  = (Resolve-Path -LiteralPath $LeftRoot).Path
$RightRoot = (Resolve-Path -LiteralPath $RightRoot).Path

if (-not $DiffOutputDir) {
  $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
  if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
  $DiffOutputDir = Join-Path $scriptDir "_diffs"
}
if ($WriteDiffs -and -not (Test-Path $DiffOutputDir)) {
  New-Item -ItemType Directory -Path $DiffOutputDir | Out-Null
}

Write-Verbose "Scanning trees..."
$leftFiles  = Get-TextFiles -Root $LeftRoot -Exts $Extensions -Exclude $ExcludePattern
$rightFiles = Get-TextFiles -Root $RightRoot -Exts $Extensions -Exclude $ExcludePattern

# Index by relative path
$leftMap  = @{}
$rightMap = @{}
foreach ($f in $leftFiles)  { $leftMap[$f.RelativePath]  = $f }
foreach ($f in $rightFiles) { $rightMap[$f.RelativePath] = $f }

$allRel = ($leftMap.Keys + $rightMap.Keys) | Select-Object -Unique | Sort-Object

$results = New-Object System.Collections.Generic.List[object]
$idx = 0
$total = $allRel.Count

foreach ($rel in $allRel) {
  $idx++
  $pct = if ($total -gt 0) { ($idx / $total) * 100 } else { 100 }
  Write-Progress -Activity "Comparing files" -Status $rel -PercentComplete $pct

  $l = $leftMap[$rel]
  $r = $rightMap[$rel]

  if ($l -and -not $r) {
    $results.Add([PSCustomObject]@{
      Status='OnlyInLeft'; RelativePath=$rel
      LeftFullPath=$l.FullPath; RightFullPath=$null
      LeftSize=$l.Length; RightSize=$null
      LeftHash=$null; RightHash=$null
      DifferenceCount=$null; DiffSample=$null
      DiffFile=$null; DiffLeftFile=$null; DiffRightFile=$null
    })
    continue
  }
  if ($r -and -not $l) {
    $results.Add([PSCustomObject]@{
      Status='OnlyInRight'; RelativePath=$rel
      LeftFullPath=$null; RightFullPath=$r.FullPath
      LeftSize=$null; RightSize=$r.Length
      LeftHash=$null; RightHash=$null
      DifferenceCount=$null; DiffSample=$null
      DiffFile=$null; DiffLeftFile=$null; DiffRightFile=$null
    })
    continue
  }

  # Both exist: determine if changed
  $leftSize  = $l.Length
  $rightSize = $r.Length
  $sameByQuick = $false
  $leftHash = $null
  $rightHash = $null

  if ($Fast) {
    if ($leftSize -eq $rightSize) {
      $leftHash  = Get-FileHashSafe -Path $l.FullPath
      $rightHash = Get-FileHashSafe -Path $r.FullPath
      if ($leftHash -and $leftHash -eq $rightHash) { $sameByQuick = $true }
    }
  }

  if ($Fast -and $sameByQuick) {
    $results.Add([PSCustomObject]@{
      Status='Same'; RelativePath=$rel
      LeftFullPath=$l.FullPath; RightFullPath=$r.FullPath
      LeftSize=$leftSize; RightSize=$rightSize
      LeftHash=$leftHash; RightHash=$rightHash
      DifferenceCount=0; DiffSample=$null
      DiffFile=$null; DiffLeftFile=$null; DiffRightFile=$null
    })
    continue
  }

  # Read lines and compare
  $leftLines  = Read-LinesSafe -Path $l.FullPath -Enc $Encoding
  $rightLines = Read-LinesSafe -Path $r.FullPath -Enc $Encoding

  # Quick equal check (position-wise)
  $equal = $false
  if ($leftLines.Count -eq $rightLines.Count) {
    $equal = $true
    for ($i=0; $i -lt $leftLines.Count; $i++) {
      if ($leftLines[$i] -ne $rightLines[$i]) { $equal = $false; break }
    }
  }

  if ($equal) {
    if ($Fast -and -not $leftHash) {
      $leftHash  = Get-FileHashSafe -Path $l.FullPath
      $rightHash = Get-FileHashSafe -Path $r.FullPath
    }
    $results.Add([PSCustomObject]@{
      Status='Same'; RelativePath=$rel
      LeftFullPath=$l.FullPath; RightFullPath=$r.FullPath
      LeftSize=$leftSize; RightSize=$rightSize
      LeftHash=$leftHash; RightHash=$rightHash
      DifferenceCount=0; DiffSample=$null
      DiffFile=$null; DiffLeftFile=$null; DiffRightFile=$null
    })
  } else {
    # Count differences line-by-line and capture a few samples
    $max = [Math]::Max($leftLines.Count, $rightLines.Count)
    $diffCount = 0
    $samples = New-Object System.Collections.Generic.List[string]
    for ($i=0; $i -lt $max; $i++) {
      $lLine = if ($i -lt $leftLines.Count) { $leftLines[$i] } else { $null }
      $rLine = if ($i -lt $rightLines.Count) { $rightLines[$i] } else { $null }
      if ($lLine -ne $rLine) {
        $diffCount++
        if ($samples.Count -lt 3) {
          $samples.Add("Line $($i+1): LEFT='${([string]$lLine)}' | RIGHT='${([string]$rLine)}'")
        }
      }
    }

    $diffPath = $null
    $leftOut = $null
    $rightOut = $null

    if ($WriteDiffs) {
      $safeName = ($rel -replace '[:\\/*?\"<>|]', '_')

      # Always produce split outputs when -WriteDiffs is set
      $leftOut  = Join-Path $DiffOutputDir ($safeName + ".left.txt")
      $rightOut = Join-Path $DiffOutputDir ($safeName + ".right.txt")
      $leftLines  | Set-Content -LiteralPath $leftOut  -Encoding UTF8
      $rightLines | Set-Content -LiteralPath $rightOut -Encoding UTF8

      # Only produce combined unified diff when NOT SplitDiffsOnly
      if (-not $SplitDiffsOnly) {
        $diffPath = Join-Path $DiffOutputDir ($safeName + ".diff.txt")
        Write-UnifiedDiff -LeftLines $leftLines -RightLines $rightLines `
          -LeftLabel ("a/"+$rel) -RightLabel ("b/"+$rel) -OutFile $diffPath
      }
    }

    if ($Fast -and -not $leftHash) {
      $leftHash  = Get-FileHashSafe -Path $l.FullPath
      $rightHash = Get-FileHashSafe -Path $r.FullPath
    }

    $results.Add([PSCustomObject]@{
      Status='Different'; RelativePath=$rel
      LeftFullPath=$l.FullPath; RightFullPath=$r.FullPath
      LeftSize=$leftSize; RightSize=$rightSize
      LeftHash=$leftHash; RightHash=$rightHash
      DifferenceCount=$diffCount
      DiffSample=($samples -join [Environment]::NewLine)
      DiffFile=$diffPath
      DiffLeftFile=$leftOut
      DiffRightFile=$rightOut
    })
  }
}

# Output & optional CSV
$results = $results | Sort-Object Status, RelativePath
$results

if ($ReportCsv) {
  try {
    $results | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $ReportCsv
    Write-Host "CSV written: $ReportCsv"
  } catch {
    Write-Warning "Failed to write CSV '$ReportCsv': $($_.Exception.Message)"
  }
}

# Summary
$summary = $results | Group-Object Status | Sort-Object Name | ForEach-Object { "{0}: {1}" -f $_.Name, $_.Count }
Write-Host ""
Write-Host "Summary"
Write-Host "-------"
$summary | ForEach-Object { Write-Host $_ }

if ($WriteDiffs) {
  if ($SplitDiffsOnly) {
    Write-Host "`nPer-file LEFT/RIGHT outputs saved under: $DiffOutputDir"
  } else {
    Write-Host "`nPer-file diffs (plus LEFT/RIGHT) saved under: $DiffOutputDir"
  }
}

# Auto-open a diff viewer if requested and we have split outputs
if ($OpenDiffViewer -and $WriteDiffs) {
  $viewer = Resolve-DiffViewer -Preference $DiffViewer
  $changed = $results | Where-Object { $_.Status -eq 'Different' -and $_.DiffLeftFile -and $_.DiffRightFile }

  if (-not $viewer) {
    Write-Warning "No supported diff viewer found. Install Visual Studio (2019/2022) or VS Code, or set -DiffViewer VSCode/VisualStudio."
  } elseif (-not $changed) {
    Write-Host "No changed files to open in the viewer."
  } else {
    $toOpen = $changed | Select-Object -First ([Math]::Max(0,[Math]::Min($MaxOpen, [int]$changed.Count)))
    Write-Host ("Opening {0} diff(s) in {1}..." -f $toOpen.Count, $viewer.Name)
    foreach ($c in $toOpen) {
      Open-Diff -ViewerName $viewer.Name -ViewerPath $viewer.Path -LeftFile $c.DiffLeftFile -RightFile $c.DiffRightFile
    }
    if ($changed.Count -gt $toOpen.Count) {
      Write-Host ("(Skipped {0} additional diff(s); increase -MaxOpen to open more.)" -f ($changed.Count - $toOpen.Count))
    }
  }
}
