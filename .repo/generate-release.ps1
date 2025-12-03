<#
.SYNOPSIS
Generates a release note markdown from manifest.json and template.

QUICK GUIDE:
1. Update manifest.json with the new version number (file-version or product-version)
2. Build your .exe and place it in the project directory (e.g., sailor-events/Sailor Events.exe)
3. Run: .\.repo\generate-release.ps1 -ProjectId "Sailor Events"
   - This generates the release notes in .repo/releases/<project-slug>/<tag>.md
4. Review the generated markdown file
5. To publish to GitHub: .\.repo\generate-release.ps1 -ProjectId "Sailor Events" -Publish
   - Creates git tag (e.g., sailor-events-v2.0.0.0)
   - Creates GitHub release
   - Uploads .exe files from project directory

Optional flags:
  -DryRun    : Show what would happen without executing
  -Latest    : Also create/update a <project-slug>-latest tag with stable download URL
  -SeriesTag : Create/update a <project-slug> tag for filtering all releases

.PARAMETER ProjectId
The project name from manifest.json (e.g., 'Sailor Events').

.PARAMETER Output
Optional output path for the generated markdown. Defaults to .repo/releases/<project-slug>/<tag>.md

.PARAMETER Publish
If provided, also creates/updates a GitHub Release for the computed tag and uploads assets.

.PARAMETER DryRun
Simulate publish actions (print commands) without executing them.

.EXAMPLE
.\.repo\generate-release.ps1 -ProjectId "Sailor Events"
.\.repo\generate-release.ps1 -ProjectId "SRO Sleep" -Publish -Latest
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$ProjectId,
    [string]$Output,
    [switch]$Publish,
    [switch]$DryRun,
    [string]$Ref = 'HEAD',
    [switch]$ForceTag,
    [switch]$SeriesTag,
    [switch]$Latest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSCommandPath | Split-Path -Parent

# Read manifest.json from repo root
$manifestPath = Join-Path $root 'manifest.json'
if (-not (Test-Path $manifestPath)) {
    throw "Manifest not found at: $manifestPath"
}

$manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json

# Get project from manifest (keyed by project name)
$project = $manifest.$ProjectId
if (-not $project) { throw "Project '$ProjectId' not found in manifest." }

# Get repo info from git
$repoOwner = $null
$repoName = $null
$defaultBranch = 'main'

try {
    $remoteUrl = (git remote get-url origin 2>$null).Trim()
    if ($remoteUrl -match 'github\.com[:/]([^/]+)/([^/\.]+)') {
        $repoOwner = $matches[1]
        $repoName = $matches[2]
    }
    $defaultBranch = (git symbolic-ref refs/remotes/origin/HEAD 2>$null).Trim() -replace '^refs/remotes/origin/', ''
    if (-not $defaultBranch) { $defaultBranch = 'main' }
} catch {
    Write-Warning "Could not detect git repo info. Using defaults."
}

# Auto-derive project slug from project name (e.g., "Sailor Events" -> "sailor-events")
$projectSlug = $ProjectId.ToLower() -replace '\s+', '-'
$projectDir = $projectSlug

# Get version from manifest
$version = $null
if ($project.PSObject.Properties.Name -contains 'file-version' -and $project.'file-version') {
    $version = [string]$project.'file-version'
} elseif ($project.PSObject.Properties.Name -contains 'product-version' -and $project.'product-version') {
    $version = [string]$project.'product-version'
}
if (-not $version) {
    throw "Version not found in manifest for project '$ProjectId'. Add 'file-version' or 'product-version' field."
}

# Auto-generate tag (format: <project-slug>-v<version>)
$tag = "$projectSlug-v$version"

# Auto-generate image URL
$imageRelPath = ".repo/resources/$projectSlug/card.png"
$imageUrl = "https://raw.githubusercontent.com/$repoOwner/$repoName/$defaultBranch/$imageRelPath"

# Load template
$templatePath = Join-Path $root '.repo/release-template.md'
if (-not (Test-Path $templatePath)) { throw "Template not found: $templatePath" }
$template = Get-Content $templatePath -Raw

$releaseDate = (Get-Date).ToString('yyyy-MM-dd')

# Get project display name
$projectName = if ($project.PSObject.Properties.Name -contains 'product-name' -and $project.'product-name') {
    [string]$project.'product-name'
} else {
    $ProjectId
}

# Extract additional fields from manifest
$notes = if ($project.PSObject.Properties.Name -contains 'notes' -and $project.notes) { [string]$project.notes } else { 'The author has not provided specific release notes for this version.' }
$companyName = if ($project.PSObject.Properties.Name -contains 'company-name' -and $project.'company-name') { [string]$project.'company-name' } else { '' }
$copyright = if ($project.PSObject.Properties.Name -contains 'copyright' -and $project.copyright) { [string]$project.copyright } else { '' }

# Simple variable replacement (literal string replace)
$content = $template
$replacements = @{
    '{{projectName}}' = $projectName
    '{{version}}'     = $version
    '{{tag}}'         = $tag
    '{{imageUrl}}'    = $imageUrl
    '{{releaseDate}}' = $releaseDate
    '{{notes}}'       = $notes
    '{{companyName}}' = $companyName
    '{{copyright}}'   = $copyright
}

# Default asset extension is .exe
$assetExtLabel = '.exe'
$replacements['{{assetExt}}'] = $assetExtLabel
foreach ($key in $replacements.Keys) {
    $content = $content.Replace($key, [string]$replacements[$key])
}

# Output path: .repo/releases/<project-slug>/<tag>.md
if (-not $Output) {
    $Output = Join-Path $root ".repo/releases/$projectSlug/$tag.md"
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Output) | Out-Null
Set-Content -Path $Output -Value $content -Encoding UTF8

Write-Host "Release note generated: $Output"

if ($Publish) {
    Write-Host "Publishing GitHub release for tag '$tag'..."

    function Run-Cmd {
        param(
            [string]$Exe,
            [Parameter(ValueFromRemainingArguments=$true)]
            [string[]]$CmdArgs
        )
        $argsStr = ($CmdArgs | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }) -join ' '
        Write-Host "> $Exe $argsStr"
        if ($DryRun) { return }
        & $Exe @CmdArgs
        if ($LASTEXITCODE -ne 0) { throw "$Exe exited with code $LASTEXITCODE" }
    }

    # Resolve executables (skip checks in DryRun)
    $gitExe = 'git'
    $ghExe = 'gh'
    if (-not $DryRun) {
        $gitCmd = Get-Command git -ErrorAction SilentlyContinue
        if (-not $gitCmd) { throw "git is not available in PATH. Install Git and retry." }
        $gitExe = $gitCmd.Source

        $ghCmd = Get-Command gh -ErrorAction SilentlyContinue
        if ($ghCmd) {
            $ghExe = $ghCmd.Source
        } else {
            $ghCandidates = @(
                (Join-Path $env:ProgramFiles 'GitHub CLI/gh.exe'),
                (Join-Path $env:LOCALAPPDATA 'Programs/GitHub CLI/gh.exe')
            )
            foreach ($cand in $ghCandidates) { if ($cand -and (Test-Path $cand)) { $ghExe = $cand; break } }
            if (-not (Test-Path $ghExe)) {
                throw 'gh (GitHub CLI) is not available in PATH. Install it or run: & "$env:ProgramFiles\GitHub CLI\gh.exe" auth login'
            }
        }
    }

    # Determine if release exists (only in real run)
    $releaseExists = $false
    if (-not $DryRun) {
        try { & $ghExe release view $tag | Out-Null; if ($LASTEXITCODE -eq 0) { $releaseExists = $true } } catch { $releaseExists = $false }
    }

    $title = "$projectName v$version"

    # Resolve desired ref commit
    $desiredSha = $Ref
    if (-not $DryRun) {
        $desiredSha = (& $gitExe rev-parse $Ref).Trim()
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($desiredSha)) {
            throw "Failed to resolve ref '$Ref' to a commit."
        }
    }

    # Ensure tag exists locally at desired ref
    $forcePush = $false
    if (-not $DryRun) {
        $existingTag = (& $gitExe tag --list $tag) | Where-Object { $_ -eq $tag }
        if (-not $existingTag) {
            Run-Cmd $gitExe 'tag' $tag $desiredSha
        } else {
            $existingSha = $null
            try { $existingSha = (& $gitExe rev-parse $tag).Trim() } catch { $existingSha = $null }
            if ($existingSha -and ($existingSha -ne $desiredSha)) {
                if ($ForceTag) {
                    Write-Host "Retagging '$tag' from $existingSha to $desiredSha ..."
                    Run-Cmd $gitExe 'tag' '-f' $tag $desiredSha
                    $forcePush = $true
                } else {
                    Write-Warning "Tag '$tag' already points to $existingSha, which differs from ref '$Ref' ($desiredSha). Use -ForceTag to retag, or specify -Ref 'main'."
                }
            } else {
                Write-Host "Tag '$tag' already exists at $existingSha."
            }
        }
    } else {
        Run-Cmd $gitExe 'tag' $tag $desiredSha
    }

    # Push tag (force if retagged)
    if ($forcePush) { Run-Cmd $gitExe 'push' '--force' 'origin' $tag } else { Run-Cmd $gitExe 'push' 'origin' $tag }

    # Optionally move/update a moving series tag (e.g., 'sailor-events') for easy filtering
    if ($SeriesTag) {
        $series = $projectSlug
        if (-not $DryRun) {
            $existingSeries = (& $gitExe tag --list $series) | Where-Object { $_ -eq $series }
            if (-not $existingSeries) {
                Run-Cmd $gitExe 'tag' $series $desiredSha
            } else {
                Run-Cmd $gitExe 'tag' '-f' $series $desiredSha
            }
        } else {
            Run-Cmd $gitExe 'tag' $series $desiredSha
        }
        Run-Cmd $gitExe 'push' '--force' 'origin' $series
    }

    if ($releaseExists -and -not $DryRun) {
           Run-Cmd $ghExe 'release' 'edit' $tag '--title' $title '--notes-file' $Output
    } else {
           Run-Cmd $ghExe 'release' 'create' $tag '--title' $title '--notes-file' $Output
    }

    # Gather .exe files from project directory
    $assetFiles = @()
    $found = Get-ChildItem -Path (Join-Path $root $projectSlug) -Filter "*.exe" -File -ErrorAction SilentlyContinue
    if ($found) { $assetFiles += $found.FullName }

    if ($assetFiles.Count -gt 0) {
        Write-Host ("Uploading assets (" + ($assetFiles -join ', ') + ")")
        Run-Cmd $ghExe 'release' 'upload' $tag @assetFiles '--clobber'
    } else {
        Write-Warning "No .exe files found in '$projectSlug'. Skipping upload."
    }

    # Optionally move/update a per-app latest tag and create/update a matching release with stable asset name
    if ($Latest) {
        $latestTagName = "$projectSlug-latest"

        # Create or force-update the lightweight tag locally
        if (-not $DryRun) {
            $existingLatest = (& $gitExe tag --list $latestTagName) | Where-Object { $_ -eq $latestTagName }
            if (-not $existingLatest) {
                Run-Cmd $gitExe 'tag' $latestTagName $desiredSha
            } else {
                Run-Cmd $gitExe 'tag' '-f' $latestTagName $desiredSha
            }
        } else {
            Run-Cmd $gitExe 'tag' $latestTagName $desiredSha
        }
        Run-Cmd $gitExe 'push' '--force' 'origin' $latestTagName

        # Ensure a GitHub Release exists for the latest tag (so resources have a stable URL)
        $latestReleaseExists = $false
        if (-not $DryRun) {
            try { & $ghExe release view $latestTagName | Out-Null; if ($LASTEXITCODE -eq 0) { $latestReleaseExists = $true } } catch { $latestReleaseExists = $false }
        }

    $latestTitle = "$projectName v$version"
        if ($latestReleaseExists -and -not $DryRun) {
            Run-Cmd $ghExe 'release' 'edit' $latestTagName '--title' $latestTitle '--notes-file' $Output
        } else {
            Run-Cmd $ghExe 'release' 'create' $latestTagName '--title' $latestTitle '--notes-file' $Output
        }

        # Upload assets to the latest release with a stable name
        if ($assetFiles.Count -gt 0) {
            $assetArgs = @()
            if ($assetFiles.Count -eq 1) {
                $ext = [System.IO.Path]::GetExtension($assetFiles[0])
                $stableName = "$projectSlug$ext"
                $assetArgs += ("{0}#{1}" -f $assetFiles[0], $stableName)
            } else {
                $assetArgs += $assetFiles
            }
            Run-Cmd $ghExe 'release' 'upload' $latestTagName @assetArgs '--clobber'
        } else {
            Write-Warning "No assets found to upload to latest release '$latestTagName'."
        }
    }

    Write-Host "Publish complete for '$tag'."
}
