<# 
.\Compare-TextTrees.ps1 -LeftRoot "C:\A" -RightRoot "C:\B" -Extensions .txt,.csv,.log,.json,.xml -WriteDiffs -ReportCsv ".\diff-summary.csv"

.SYNOPSIS
  Compare text files in two folder trees.

.PARAMETER LeftRoot
  Left/top folder path.

.PARAMETER RightRoot
  Right/top folder path.

.PARAMETER Extensions
  One or more file extensions to treat as "text". Defaults to .txt,.csv,.log,.json,.xml,.ps1,.psm1,.psd1,.ini,.cfg,.conf,.sql

.PARAMETER ExcludePattern
  Wildcard pattern(s) (array) to exclude relative paths, e.g. '*\bin\*','*.tmp'

.PARAMETER Encoding
  Text encoding hint for reading files. Default: utf8. (If reading fails, falls back to default).

.PARAMETER WriteDiffs
  If set, writes unified-style line diffs for changed files under a _diffs folder next to the script (or in -DiffOutputDir).

.PARAMETER DiffOutputDir
  Optional path where per-file diff text files are written.

.PARAMETER ReportCsv
  Optional path to write a CSV summary of results.

.PARAMETER Fast
  If set, uses file size + SHA256 hashes for change detection (skips line-by-line diff unless -WriteDiffs is set).
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
  [string]$DiffOutputDir,
  [string]$ReportCsv,

  [switch]$Fast
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
  $extSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($e in $Exts) { [void]$extSet.Add(($e.StartsWith('.') ? $e : ".$e")) }

  $files = Get-ChildItem -Path $Root -Recurse -File -ErrorAction Stop |
    Where-Object { $extSet.Contains([System.IO.Path]::GetExtension($_.FullName)) } |
    ForEach-Object {
      $rel = Resolve-RelativePath -Root $Root -FullPath $_.FullName
      # exclude?
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
    # Get-Content with -Encoding for PS5 compatibility
    Get-Content -LiteralPath $Path -Encoding $Enc -ErrorAction Stop
  } catch {
    # Fallback to default (auto)
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
  # Simple unified diff (context-less): prefix '-' for left-only, '+' for right-only, ' ' for same
  # Use Compare-Object to compute differences by line; not perfect for moves but adequate for quick review.
  # For more detailed diffs, additional libraries would be needed.
  $max = [Math]::Max($LeftLines.Count, $RightLines.Count)
  $out = New-Object System.Collections.Generic.List[string]
  $out.Add("--- $LeftLabel")
  $out.Add("+++ $RightLabel")
  for ($i=0; $i -lt $max; $i++) {
    $l = ($(if ($i -lt $LeftLines.Count) { $LeftLines[$i] } else { $null }))
    $r = ($(if ($i -lt $RightLines.Count) { $RightLines[$i] } else { $null }))
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

# Prep
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
  Write-Progress -Activity "Comparing files" -Status $rel -PercentComplete (($idx / [Math]::Max($total,1))*100)

  $l = $leftMap[$rel]
  $r = $rightMap[$rel]

  if ($l -and -not $r) {
    $results.Add([PSCustomObject]@{
      Status='OnlyInLeft'; RelativePath=$rel
      LeftFullPath=$l.FullPath; RightFullPath=$null
      LeftSize=$l.Length; RightSize=$null
      LeftHash=$null; RightHash=$null
      DifferenceCount=$null; DiffSample=$null
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
    })
    continue
  }

  # Read lines and compare
  $leftLines  = Read-LinesSafe -Path $l.FullPath -Enc $Encoding
  $rightLines = Read-LinesSafe -Path $r.FullPath -Enc $Encoding

  # Quick equal check
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
    })
  } else {
    # Count differences line-by-line (position-wise)
    $max = [Math]::Max($leftLines.Count, $rightLines.Count)
    $diffCount = 0
    $samples = New-Object System.Collections.Generic.List[string]
    for ($i=0; $i -lt $max; $i++) {
      $lLine = ($(if ($i -lt $leftLines.Count) { $leftLines[$i] } else { $null }))
      $rLine = ($(if ($i -lt $rightLines.Count) { $rightLines[$i] } else { $null }))
      if ($lLine -ne $rLine) {
        $diffCount++
        if ($samples.Count -lt 3) {
          $samples.Add("Line {0}: LEFT='{1}' | RIGHT='{2}'" -f ($i+1), $lLine, $rLine)
        }
      }
    }

    $diffPath = $null
    if ($WriteDiffs) {
      $safeName = ($rel -replace '[:\\/*?\"<>|]', '_')
      $diffPath = Join-Path $DiffOutputDir ($safeName + ".diff.txt")
      Write-UnifiedDiff -LeftLines $leftLines -RightLines $rightLines `
        -LeftLabel ("a/"+$rel) -RightLabel ("b/"+$rel) -OutFile $diffPath
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
if ($WriteDiffs) { Write-Host "`nPer-file diffs (if any) saved under: $DiffOutputDir" }
