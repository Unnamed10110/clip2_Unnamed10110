# Creates a GitHub release for this repo and uploads the built clip2.exe from build/.
#
# Usage:
#   .\release_github.ps1 -Version 1.5
#   .\release_github.ps1 -Version v1.5 -Notes "Fixes and paste improvements"
#   .\release_github.ps1 -Version 1.5 -Draft
#   .\release_github.ps1 -Version 1.5 -Replace
#
# Requires: GitHub CLI (`gh`) authenticated against this repo (gh auth login).

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Version,

    [string]$Notes = "",

    [string]$Title = "",

    [switch]$Draft,

    [switch]$Prerelease,

    [switch]$Replace
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }

function Find-Clip2Exe {
    $candidates = @(
        (Join-Path $root "build\clip2.exe"),
        (Join-Path $root "build\Release\clip2.exe")
    )
    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) { return (Resolve-Path -LiteralPath $path).Path }
    }
    return $null
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) is not on PATH. Install it from https://cli.github.com/ and run 'gh auth login'."
}

$exe = Find-Clip2Exe
if (-not $exe) {
    throw "clip2.exe not found under build\ or build\Release\. Run build.bat first."
}

$tag = $Version.Trim()
if ($tag -notmatch '^[vV]') { $tag = "v$tag" }
if ($tag -notmatch '^[vV][\w.+-]+$') {
    throw "Invalid version '$Version'. Use something like 1.5, v1.5, or V09.08.2026."
}

if (-not $Title) { $titleText = "clip2 $tag" } else { $titleText = $Title }
if (-not $Notes) { $notesText = "clip2 $tag" } else { $notesText = $Notes }

Write-Host "Release tag : $tag"
Write-Host "Asset       : $exe"
Write-Host "Title       : $titleText"

$existing = gh release view $tag --json tagName 2>$null
if ($LASTEXITCODE -eq 0 -and $existing) {
    if (-not $Replace) {
        throw "Release '$tag' already exists. Pass -Replace to delete it and recreate."
    }
    Write-Host "Deleting existing release $tag ..."
    gh release delete $tag --yes
    if ($LASTEXITCODE -ne 0) { throw "Failed to delete existing release '$tag'." }
}

$ghArgs = @(
    "release", "create", $tag, $exe,
    "--title", $titleText,
    "--notes", $notesText
)
if ($Draft) { $ghArgs += "--draft" }
if ($Prerelease) { $ghArgs += "--prerelease" }

Write-Host "Creating GitHub release..."
& gh @ghArgs
if ($LASTEXITCODE -ne 0) {
    throw "gh release create failed. If you are not logged in, run 'gh auth login'."
}

Write-Host ""
Write-Host "Published $tag with $(Split-Path $exe -Leaf)."
