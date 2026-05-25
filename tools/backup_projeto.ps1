param(
  [string]$ProjectRoot = "",
  [string]$BackupsDir = "",
  [switch]$WipCommit,
  [switch]$ZipOnly,
  [switch]$GitOnly
)

$ErrorActionPreference = 'Stop'

function Resolve-ProjectRoot {
  param([string]$ProjectRoot)

  if ($ProjectRoot -and (Test-Path -LiteralPath $ProjectRoot)) {
    return (Resolve-Path -LiteralPath $ProjectRoot).Path
  }

  # Default: parent folder of this script (tools\..)
  $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
  $root = Join-Path $scriptDir ".."
  return (Resolve-Path -LiteralPath $root).Path
}

function Ensure-Dir {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

function Get-RelativePath {
  param(
    [Parameter(Mandatory=$true)][string]$BasePath,
    [Parameter(Mandatory=$true)][string]$FullPath
  )
  $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\') + '\'
  $full = [System.IO.Path]::GetFullPath($FullPath)
  if ($full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
    return $full.Substring($base.Length)
  }
  return (Split-Path -Leaf $full)
}

function Copy-FileToStage {
  param(
    [Parameter(Mandatory=$true)][string]$ProjectRoot,
    [Parameter(Mandatory=$true)][string]$FullPath,
    [Parameter(Mandatory=$true)][string]$StageRoot
  )
  $rel = Get-RelativePath -BasePath $ProjectRoot -FullPath $FullPath
  $dest = Join-Path $StageRoot $rel
  $destDir = Split-Path -Parent $dest
  Ensure-Dir -Path $destDir
  Copy-Item -LiteralPath $FullPath -Destination $dest -Force
}

$root = Resolve-ProjectRoot -ProjectRoot $ProjectRoot
if (-not $BackupsDir) {
  $BackupsDir = Join-Path $root 'backups'
}
Ensure-Dir -Path $BackupsDir

$ts = (Get-Date).ToString('yyyyMMdd-HHmmss')
$stage = Join-Path $BackupsDir ("_stage_$ts")
Ensure-Dir -Path $stage

Write-Host "ProjectRoot: $root"
Write-Host "BackupsDir : $BackupsDir"
Write-Host "Timestamp  : $ts"

# --- Build include list (source + templates + build assets) ---
$includeExtensions = @(
  '.py', '.html', '.spec', '.bat', '.md', '.txt', '.json'
)

$includeExactFiles = @(
  'Logo.png',
  'splash.png',
  'mapa_brasil_dados.pkl'
)

$includeDirs = @(
  'geo_data'
)

$excludeDirs = @(
  '.git', 'venv_stable', '.venv', 'build', 'dist', '__pycache__', 'backups'
)

$excludeNamePatterns = @(
  'dados_consolidados_*.xlsx',
  '*.xlsx',
  '*.pdf',
  'grafico_*.png'
)

function Is-ExcludedPath {
  param([string]$FullPath)
  $p = $FullPath.ToLowerInvariant()
  foreach ($d in $excludeDirs) {
    $seg = "\\$d\\".ToLowerInvariant()
    if ($p.Contains($seg)) { return $true }
  }
  return $false
}

function Is-ExcludedName {
  param([string]$Name)
  foreach ($pat in $excludeNamePatterns) {
    if ($Name -like $pat) { return $true }
  }
  return $false
}

$files = New-Object System.Collections.Generic.List[string]

# collect by extensions
Get-ChildItem -LiteralPath $root -Recurse -File | ForEach-Object {
  if (Is-ExcludedPath -FullPath $_.FullName) { return }
  if (Is-ExcludedName -Name $_.Name) { return }
  if ($includeExtensions -contains $_.Extension) {
    $files.Add($_.FullName) | Out-Null
  }
}

# include exact files if exist
foreach ($name in $includeExactFiles) {
  $p = Join-Path $root $name
  if (Test-Path -LiteralPath $p) {
    $files.Add((Resolve-Path -LiteralPath $p).Path) | Out-Null
  }
}

# include directories (copy everything except exclusions)
foreach ($d in $includeDirs) {
  $dirPath = Join-Path $root $d
  if (Test-Path -LiteralPath $dirPath) {
    Get-ChildItem -LiteralPath $dirPath -Recurse -File | ForEach-Object {
      if (Is-ExcludedName -Name $_.Name) { return }
      $files.Add($_.FullName) | Out-Null
    }
  }
}

# de-duplicate
$files = $files | Sort-Object -Unique

# Write git info (even if not a git repo)
$gitInfoPath = Join-Path $stage 'git_info.txt'
$gitInfo = New-Object System.Collections.Generic.List[string]
$gitInfo.Add("Backup timestamp: $ts") | Out-Null
$gitInfo.Add("Project root    : $root") | Out-Null

Push-Location $root
try {
  $isGit = $false
  try {
    $inside = (git rev-parse --is-inside-work-tree 2>$null)
    if ($inside -eq 'true') { $isGit = $true }
  } catch { $isGit = $false }

  if ($isGit) {
    $head = (git rev-parse HEAD)
    $branch = (git branch --show-current)
    $status = (git status --porcelain)

    $gitInfo.Add("Git branch      : $branch") | Out-Null
    $gitInfo.Add("Git HEAD        : $head") | Out-Null
    $gitInfo.Add("Git dirty       : " + ([string]::IsNullOrWhiteSpace($status) -eq $false)) | Out-Null
    $gitInfo.Add("Git status:") | Out-Null
    if ($status) { $gitInfo.Add($status) | Out-Null } else { $gitInfo.Add("(clean)") | Out-Null }
  } else {
    $gitInfo.Add("Git            : not a repo") | Out-Null
  }

  $gitInfo | Out-File -FilePath $gitInfoPath -Encoding UTF8
} finally {
  Pop-Location
}

# Export settings if exist
$settingsDir = Join-Path $env:APPDATA 'relatorios_2026'
$settingsPath = Join-Path $settingsDir 'settings.json'
if (Test-Path -LiteralPath $settingsPath) {
  $dest = Join-Path $stage 'settings.json'
  Copy-Item -LiteralPath $settingsPath -Destination $dest -Force
}

# Stage files
foreach ($f in $files) {
  Copy-FileToStage -ProjectRoot $root -FullPath $f -StageRoot $stage
}

$zipPath = Join-Path $BackupsDir ("backup_projeto_$ts.zip")

if (-not $GitOnly) {
  Write-Host "Creating ZIP: $zipPath"
  Compress-Archive -LiteralPath (Join-Path $stage '*') -DestinationPath $zipPath -Force
}

# Git tag / optional WIP commit
if (-not $ZipOnly) {
  Push-Location $root
  try {
    $inside = $false
    try {
      $inside = ((git rev-parse --is-inside-work-tree) -eq 'true')
    } catch { $inside = $false }

    if ($inside) {
      $tag = "backup-$ts"
      $branch = (git branch --show-current)
      $status = (git status --porcelain)
      $dirty = -not [string]::IsNullOrWhiteSpace($status)

      if ($WipCommit -and $dirty) {
        $backupBranch = "backup/wip-$ts"
        Write-Host "Creating WIP backup branch: $backupBranch"
        git switch -c $backupBranch | Out-Null
        git add -A | Out-Null
        git commit -m "WIP backup $ts" | Out-Null
        $tag = "backup-wip-$ts"
        Write-Host "Creating git tag: $tag"
        git tag -a $tag -m "WIP backup $ts" | Out-Null
        Write-Host "Returning to branch: $branch"
        git switch $branch | Out-Null
      } else {
        Write-Host "Creating git tag (HEAD): $tag"
        git tag -a $tag -m "Backup tag $ts" | Out-Null
        if ($dirty) {
          Write-Host "NOTE: Working tree is dirty; tag points to last commit only. ZIP captures current files." -ForegroundColor Yellow
        }
      }
    } else {
      Write-Host "Git repo not detected; skipping git tag/branch." -ForegroundColor Yellow
    }
  } finally {
    Pop-Location
  }
}

# Cleanup stage
try {
  Remove-Item -LiteralPath $stage -Recurse -Force
} catch {
  Write-Host "Warning: could not remove staging folder: $stage" -ForegroundColor Yellow
}

Write-Host "Done."