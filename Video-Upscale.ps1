# ============================================================
# Interactive Upscale Script (Smart Resume)
#  - Input: Cropped videos (e.g. 640x360)
#  - Output: 1080p FFV1 (lossless) + FLAC in MKV
#
# Features:
#  * GPU selection + GPU-tagged work/output dirs (gpu0/gpu1)
#  * Resume: keeps extracted frames; and if upscale is incomplete,
#    it ONLY upscales the missing frames (no 17-hour restart)
#  * Optional keep temp dirs (NoDelete)
#  * Output validation before cleanup
# ============================================================

<#
.SYNOPSIS
Interactive video upscaler pipeline using Upscayl + ffmpeg with smart resume and GPU tagging.

.DESCRIPTION
This script batch-processes source videos (default folder: "source") by:

1) Extracting frames with ffmpeg
2) Upscaling frames with Upscayl (Real-ESRGAN model)
3) Rebuilding a final video using ffmpeg with audio copied from the original source

It supports:
- GPU selection (manual or auto) and GPU-tagged work/output folders to avoid job collisions
- Smart resume: keeps extracted frames and only upscales missing frames if an upscale run was interrupted
- Optional temp cleanup control (NoDelete)
- Output validation before cleanup
- Dry-run resume check (detect incomplete jobs without doing any work)
- Output presets (codec/container/target resolution choices)

.PARAMETER InputDir
Source folder containing input videos. Default is "source".
Example: -InputDir "source-640x360"

.PARAMETER OutputPreset
Controls final output format (container + codecs + target resolution).
Common presets:
- ffv1_1080p_mkv  : Lossless FFV1 video + FLAC audio in MKV (master)
- ffv1_4k_mkv     : Lossless FFV1 video + FLAC audio in MKV (master, 4K)
- h264_1080p_mp4  : H.264 video + AAC audio in MP4 (delivery)
- hevc_1080p_mp4  : H.265/HEVC video + AAC audio in MP4 (delivery, smaller)
- prores_1080p_mov: ProRes 422 HQ + PCM audio in MOV (editing)

Default: ffv1_1080p_mkv

.PARAMETER GpuIndex
Upscayl GPU index to use. If omitted, the script can auto-select a discrete GPU
and warn if the chosen GPU is busy/locked.

.PARAMETER Resume
Resume mode:
- If extracted frames already exist, skip extraction.
- If upscaled frames are incomplete, only upscale missing frames (smart resume).

.PARAMETER NoDelete
If set, keeps working directories (frames/upscaled) after success.
If not set, work folders are removed only after a verified output is produced.

.PARAMETER DryRun
Dry-run resume check only:
- Detects unfinished projects in work directories and reports progress
- Does not run extraction/upscale/encode

.PARAMETER Force
Overwrite output files if they exist. If not specified, output will be skipped
when a valid output already exists.

.PARAMETER UpscaylPath
Path to upscayl-bin.exe. Default is ".\upscayl-bin.exe" in the script folder.

.PARAMETER Model
Upscayl model name (e.g. realesrgan-x4plus). Default: realesrgan-x4plus

.PARAMETER Scale
Upscayl scale factor. Default: 3

.PARAMETER Extensions
Comma-separated list of input extensions to process (e.g. "mp4,mov,mkv").
Default: "mp4"

.EXAMPLE
# Basic run (default input folder "source", lossless 1080p master)
.\upscale.ps1

.EXAMPLE
# Specify input folder and output preset
.\upscale.ps1 -InputDir "source-640x360" -OutputPreset ffv1_1080p_mkv

.EXAMPLE
# Run on a specific GPU
.\upscale.ps1 -GpuIndex 1 -OutputPreset ffv1_1080p_mkv

.EXAMPLE
# Resume an interrupted job (only upscale missing frames, keep temp files)
.\upscale.ps1 -Resume -NoDelete

.EXAMPLE
# Dry-run: show incomplete project progress without doing work
.\upscale.ps1 -DryRun

.EXAMPLE
# Create a delivery MP4 and overwrite existing output
.\upscale.ps1 -OutputPreset h264_1080p_mp4 -Force

.NOTES
- Requires ffmpeg + ffprobe available in PATH.
- Upscayl GPU indexing is determined by Upscayl. Use your script’s GPU listing/snapshot
  output to confirm index mappings.
- Work directories are GPU-tagged (e.g. _ffv1_work_gpu0) to avoid collisions.
- Output is validated before any cleanup occurs.

.LINK
ffmpeg: https://ffmpeg.org/
Upscayl: https://github.com/upscayl/upscayl
#>


[CmdletBinding()]
param(
    # Paths
    [Parameter()]
    [string]$InputDir = "source",

    [Parameter()]
    [string]$script:UpscaylExe = ".\upscayl-bin.exe",

    # Upscayl options
    [Parameter()]
    [ValidateSet("realesrgan-x4plus","realesrgan-x4plus-anime","remacri","ultramix_balanced","ultramix_ultrasharp")]
    [string]$Model = "realesrgan-x4plus",

    [Parameter()]
    [ValidateRange(1,8)]
    [int]$Scale = 3,

    [Parameter()]
    [ValidateRange(-1,32)]
    [int]$GpuIndex = -1,   # -1 = auto

    # Output format choice (expanded + default = best file size)
    [Parameter()]
    [ValidateSet(
        "ffv1_1080p_mkv","ffv1_4k_mkv",
        "h264_1080p_mp4","h264_4k_mp4",
        "hevc_1080p_mp4","hevc_4k_mp4",
        "prores_1080p_mov","prores_4k_mov",
        "av1_1080p_mkv","av1_1080p_mp4",
        "av1_4k_mkv","av1_4k_mp4"
    )]
    [string]$OutputPreset = "av1_1080p_mkv",

    # Output settings
    [Parameter()]
    [string]$OutputDir = "",

    [Parameter()]
    [string]$WorkDir = "",

    # Behavior switches
    [Parameter()]
    [switch]$Resume,

    [Parameter()]
    [switch]$NoDelete,

    [Parameter()]
    [switch]$AutoDownloadTools,

    [Parameter()]
    [switch]$DryRun,

    [Parameter()]
    [string]$Extensions = "",   # empty = use default video list

    [Parameter()]
    [switch]$Force,        # overwrite existing output if set

    [Parameter()]
    [ValidateRange(0,100)]
    [int]$BusyThresholdPercent = 50,

    [Parameter()]
    [switch]$Help,

    [Parameter()]
    [switch]$NoPreferDiscrete
)

$PreferDiscrete = -not $NoPreferDiscrete

if ($Help) {
    Get-Help -Detailed $PSCommandPath
    exit 0
}

# Canonical video extensions (do NOT store this in $Extensions)
$VideoExtensions = @(
  "mp4","mkv","mov","avi","webm","m4v","mpg","mpeg",
  "wmv","ts","mts","m2ts","3gp","3g2","flv","f4v",
  "vob","ogv","rm","rmvb","asf","dv","y4m"
)

if (-not $PSBoundParameters.ContainsKey("DryRun")) {
    $DryRun = (Read-Host "Dry-run resume check only? (Y/N)").ToUpper() -eq "Y"
}
function Test-IsVideoFile {
    param([Parameter(Mandatory)][string]$Path)
    try {
        $types = & $script:FfprobeExe -v error -show_entries stream=codec_type -of csv=p=0 "file:$Path"
        return ($types -match 'video')
    } catch {
        return $false
    }
}

function Get-ImageSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FfprobeExe,
        [Parameter(Mandatory)][string]$ImagePath
    )

    if (-not (Test-Path -LiteralPath $ImagePath)) {
        throw "Get-ImageSize: file not found: $ImagePath"
    }

    $ffArgs = @(
        "-v","error",
        "-select_streams","v:0",
        "-show_entries","stream=width,height",
        "-of","csv=p=0:s=x",
        $ImagePath
    )

    $out = & $FfprobeExe @ffargs 2>$null
    if (-not $out) { throw "Get-ImageSize: ffprobe returned no output for $ImagePath" }

    $parts = $out.Trim() -split "x"
    if ($parts.Count -ne 2) { throw "Get-ImageSize: unexpected ffprobe output '$out' for $ImagePath" }

    [pscustomobject]@{
        Width  = [int]$parts[0]
        Height = [int]$parts[1]
    }
}

function Get-TargetWHFromPreset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Preset,
        [Parameter(Mandatory)][pscustomobject]$SrcInfo
    )

    # If preset.Scale is null (match_source), treat target as source
    if ([string]::IsNullOrWhiteSpace($Preset.Scale)) {
        return [pscustomobject]@{ W = [int]$SrcInfo.Width; H = [int]$SrcInfo.Height; Mode = "match_source" }
    }

    # preset.Scale looks like "1920:1080"
    $parts = $Preset.Scale -split ":"
    if ($parts.Count -ne 2) {
        throw "Preset.Scale is invalid: '$($Preset.Scale)'. Expected 'W:H'."
    }

    return [pscustomobject]@{ W = [int]$parts[0]; H = [int]$parts[1]; Mode = "fixed" }
}

function Write-WarningIfNoResolutionGain {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$SrcInfo,
        [Parameter(Mandatory)][int]$TargetW,
        [Parameter(Mandatory)][int]$TargetH,
        [Parameter(Mandatory)][int]$UpscaylScale,
        [switch]$PromptToContinue
    )

    $srcW = [int]$SrcInfo.Width
    $srcH = [int]$SrcInfo.Height

    # Case A: Source already equals output target
    if ($srcW -eq $TargetW -and $srcH -eq $TargetH) {
        Write-Warning ("[Resolution] Source is already {0}x{1}, which matches the selected output target {2}x{3}." -f $srcW,$srcH,$TargetW,$TargetH)
        Write-Warning ("[Resolution] Upscaling may not improve detail much. If you upscale (x{0}) and then rebuild to the same size, you may be mostly sharpening/denoising, not adding real detail." -f $UpscaylScale)

        if ($PromptToContinue) {
            $resp = Read-Host "Continue anyway? (Y/N)"
            if ($resp.Trim().ToUpper() -ne "Y") { throw "Aborted by user (no resolution gain)." }
        }
        return
    }

    # Case B: You upscale larger than target and then scale DOWN to target (wasted compute)
    # Upscayl output approx = src * scale
    $upW = $srcW * $UpscaylScale
    $upH = $srcH * $UpscaylScale

    if ($upW -ge $TargetW -and $upH -ge $TargetH -and ($TargetW -le $srcW -or $TargetH -le $srcH)) {
        # This is a “probably wasted” scenario only if target is <= source in at least one dimension.
        Write-Warning ("[Resolution] Your Upscayl output would be ~{0}x{1} (x{2}), but the selected output target is {3}x{4}." -f $upW,$upH,$UpscaylScale,$TargetW,$TargetH)
        Write-Warning "[Resolution] That means you’ll upscale and then downscale during rebuild—big compute cost for limited benefit."

        if ($PromptToContinue) {
            $resp = Read-Host "Continue anyway? (Y/N)"
            if ($resp.Trim().ToUpper() -ne "Y") { throw "Aborted by user (upscale then downscale warning)." }
        }
    }
}


function Get-SampleFramePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FramesDir
    )

    if (-not (Test-Path -LiteralPath $FramesDir)) { return $null }

    # pick the first png/jpg we find (sorted) as a sample
    $sample = Get-ChildItem -LiteralPath $FramesDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match '^\.(png|jpg|jpeg|webp)$' } |
        Sort-Object Name |
        Select-Object -First 1

    return $sample?.FullName
}

function Test-FramesMatchTarget{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FfmpegExe,
        [Parameter(Mandatory)][string]$FfprobeExe,

        [Parameter(Mandatory)][string]$UpscaledFramesDir,

        [Parameter(Mandatory)][int]$TargetW,
        [Parameter(Mandatory)][int]$TargetH,

        # Where normalized frames will go IF needed
        [Parameter(Mandatory)][string]$NormalizedFramesDir,

        # Your frame pattern like frame_%06d.png or similar
        [Parameter(Mandatory)][string]$FramePattern,

        # If your frames are PNG keep png; if JPG keep jpg
        [Parameter][ValidateSet("png","jpg")][string]$OutExt = "png",

        # Aspect policy: "Pad" = keep AR, add bars; "Crop" = fill frame; "Stretch" = force
        [ValidateSet("Pad","Crop","Stretch")]
        [string]$AspectPolicy = "Pad",

        [switch]$DryRun
    )

    $sample = Get-SampleFramePath -FramesDir $UpscaledFramesDir
    if (-not $sample) { throw "No frames found in: $UpscaledFramesDir" }

    $size = Get-ImageSize -FfprobeExe $FfprobeExe -ImagePath $sample
    Write-Host ("[Frames] Upscayl sample: {0}x{1} (target {2}x{3})" -f $size.Width,$size.Height,$TargetW,$TargetH)

    if ($size.Width -eq $TargetW -and $size.Height -eq $TargetH) {
        Write-Host "[Frames] Already matches target; skipping normalization."
        return $UpscaledFramesDir
    }

    Write-Warning "[Frames] Mismatch: will normalize with ffmpeg to exact target size."

    if ($DryRun) {
        Write-Host "[DryRun] Would normalize frames to $TargetW x $TargetH into: $NormalizedFramesDir"
        return $NormalizedFramesDir
    }

    New-Item -ItemType Directory -Force -Path $NormalizedFramesDir | Out-Null

    # Build vf per policy
    $vf = switch ($AspectPolicy) {
        "Pad" {
            # keep AR, fit inside, pad to exact
            # setsar=1 keeps square pixels
            # Pad
            "scale=w=${TargetW}:h=${TargetH}:force_original_aspect_ratio=decrease,pad=${TargetW}:${TargetH}:(ow-iw)/2:(oh-ih)/2,setsar=1"

        }
        "Crop" {
            # keep AR, fill, then crop to exact
            # Crop
            "scale=w=${TargetW}:h=${TargetH}:force_original_aspect_ratio=increase,crop=${TargetW}:${TargetH},setsar=1"
        }
        "Stretch" {
            # force exact (may distort)
            # Stretch
            "scale=w=${TargetW}:h=${TargetH},setsar=1"
        }
    }

    $inPattern  = Join-Path $UpscaledFramesDir $FramePattern
    $outPattern = Join-Path $NormalizedFramesDir ($FramePattern -replace '\.\w+$', ".$OutExt")

    Write-Host "[Frames] Normalizing with vf: $vf"

    & $FfmpegExe @(
        "-hide_banner","-loglevel","error",
        "-y",
        "-i", $inPattern,
        "-vf", $vf,
        # keep lossless-ish for PNG; adjust if you use JPG
        $outPattern
    )

    # Validate normalized sample
    $sample2 = Get-SampleFramePath -FramesDir $NormalizedFramesDir
    if (-not $sample2) { throw "Normalization produced no frames in: $NormalizedFramesDir" }

    $size2 = Get-ImageSize -FfprobeExe $FfprobeExe -ImagePath $sample2
    if ($size2.Width -ne $TargetW -or $size2.Height -ne $TargetH) {
        throw "Normalization failed: got $($size2.Width)x$($size2.Height) expected $TargetW x $TargetH"
    }

    Write-Host "[Frames] Normalization OK."
    return $NormalizedFramesDir
}


function Resolve-VideoTools {
    [CmdletBinding()]
    param(
        [string]$FfmpegExe,
        [string]$FfprobeExe,
        [string]$UpscaylExe,

        [switch]$AutoDownloadTools,

        # optional: folder where your script downloads tools
        [string]$ToolsDir = (Join-Path $PSScriptRoot "tools")
    )

    function _ResolveOne {
        param(
            [Parameter(Mandatory)][string]$Name,
            [string]$PreferredPath,
            [string]$CommandName
        )

        # 1) Explicit path wins (if supplied)
        if (-not [string]::IsNullOrWhiteSpace($PreferredPath)) {
            if (Test-Path -LiteralPath $PreferredPath) {
                return (Resolve-Path -LiteralPath $PreferredPath).Path
            }
            throw "$Name path was provided but not found: $PreferredPath"
        }

        # 2) PATH
        $cmd = Get-Command $CommandName -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source) { return $cmd.Source }

        return $null
    }

    $resolved = [ordered]@{
        ffmpeg  = _ResolveOne -Name "ffmpeg"  -PreferredPath $FfmpegExe  -CommandName "ffmpeg"
        ffprobe = _ResolveOne -Name "ffprobe" -PreferredPath $FfprobeExe -CommandName "ffprobe"
        upscayl = _ResolveOne -Name "upscayl" -PreferredPath $UpscaylExe -CommandName "upscayl"
    }

    # If you want: optionally attempt downloads ONLY if enabled and something missing
    if ($AutoDownloadTools) {
        foreach ($k in @("ffmpeg","ffprobe","upscayl")) {
            if (-not $resolved[$k]) {
                # Call your existing download logic here (keep it in ONE place)
                # $resolved[$k] = Get-DownloadedToolPath -Name $k -ToolsDir $ToolsDir
            }
        }
    }

    # Final validation (fail fast, clean error)
    foreach ($k in @("ffmpeg","ffprobe","upscayl")) {
        if (-not $resolved[$k]) {
            throw "[Tools] Missing required tool: $k. Provide -${k}Exe or add it to PATH (or enable -AutoDownloadTools)."
        }
    }

    # Export as one canonical script state
    $script:Tools = [pscustomobject]@{
        Ffmpeg  = $resolved.ffmpeg
        Ffprobe = $resolved.ffprobe
        Upscayl = $resolved.upscayl
    }

    return $script:Tools
}

function Get-InputVideoFiles {
    param(
        [Parameter(Mandatory)][string]$InputDir,
        [string]$Extensions
    )

    # "*" means: scan all files and detect video by ffprobe
    if ($Extensions -and $Extensions.Trim() -eq "*") {
        return Get-ChildItem -Path $InputDir -File -ErrorAction SilentlyContinue |
            Where-Object { Test-IsVideoFile $_.FullName } |
            Sort-Object FullName -Unique
    }

    # extension-only scanning
    $exts =
        if ([string]::IsNullOrWhiteSpace($Extensions)) {
            $VideoExtensions
        } else {
            $Extensions.Split(',') |
                ForEach-Object { $_.Trim().TrimStart('.').ToLowerInvariant() } |
                Where-Object { $_ }
        }

    return Get-ChildItem -Path $InputDir -File -ErrorAction SilentlyContinue |
        Where-Object {
            $ext = $_.Extension.TrimStart('.').ToLowerInvariant()
            $exts -contains $ext
        } |
        Sort-Object FullName -Unique
}

$null = Resolve-VideoTools `
    -FfmpegExe  $FfmpegExe `
    -FfprobeExe $FfprobeExe `
    -UpscaylExe $UpscaylExe `
    -AutoDownloadTools:$AutoDownloadTools

$BaseDir = (Get-Location).Path

# ---------------------------
# Tool discovery (PS5.1-safe)
# Priority: Param path -> Script folder -> PATH
# ---------------------------

# Script folder (works in PS5.1)
$script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Get-UpscaylScaleForTarget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$SrcW,
        [Parameter(Mandatory)][int]$SrcH,
        [Parameter(Mandatory)][int]$TargetW,
        [Parameter(Mandatory)][int]$TargetH,
        [Parameter()][int]$MaxScale = 4,
        [Parameter()][int]$MinScale = 1
    )

    # ratio needed to reach/beat target in each dimension
    $rw = $TargetW / [double]$SrcW
    $rh = $TargetH / [double]$SrcH

    # we want upscayl output >= target in BOTH dimensions (so normalization is only minor)
    $need = [math]::Ceiling([math]::Max($rw, $rh))

    # clamp to what you allow
    $need = [math]::Max($MinScale, [math]::Min($MaxScale, $need))

    return [int]$need
}
function Resolve-ExeCandidate {
    param([Parameter(Mandatory)][string]$Path)
    try {
        if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
        if (Test-Path -LiteralPath $Path) { return (Resolve-Path -LiteralPath $Path).Path }
    } catch { }
    return $null
}

function Find-ExeInDir {
    param(
        [Parameter(Mandatory)][string]$Dir,
        [Parameter(Mandatory)][string[]]$Names
    )
    foreach ($n in $Names) {
        $p = Join-Path $Dir $n
        $r = Resolve-ExeCandidate -Path $p
        if ($r) { return $r }
    }
    return $null
}

function Resolve-ExeFromPath {
    param([Parameter(Mandatory)][string]$Name)
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) { return $cmd.Source }
    return $null
}


function Resolve-FullPath {
    param([Parameter(Mandatory)][string]$Path, [string]$Base = $BaseDir)

    if ([IO.Path]::IsPathRooted($Path)) {
        return (Resolve-Path $Path -ErrorAction Stop).Path
    } else {
        return (Resolve-Path (Join-Path $Base $Path) -ErrorAction Stop).Path
    }
}

# ===========================
# GitHub Release Auto-Download
# ===========================
# Requires: PowerShell 7+, Expand-Archive, Internet access
# Uses:
#   - https://api.github.com/repos/GyanD/codexffmpeg/releases/latest
#   - https://api.github.com/repos/upscayl/upscayl/releases/latest

$ToolsRoot = Join-Path $BaseDir "tools"
New-Item -ItemType Directory -Force $ToolsRoot | Out-Null

function Invoke-GitHubApi {
    param([Parameter(Mandatory)][string]$Url)

    $headers = @{
        "User-Agent" = "Video-Upscale-Script"
        "Accept"     = "application/vnd.github+json"
    }

    # Optional: if you want to avoid rate limits, set env var GITHUB_TOKEN
    if ($env:GITHUB_TOKEN) {
        $headers["Authorization"] = "Bearer $($env:GITHUB_TOKEN)"
    }

    return Invoke-RestMethod -Uri $Url -Headers $headers -Method Get -ErrorAction Stop
}
function Get-PresetPixFmt {
    param([Parameter(Mandatory)]$Preset)

    try {
        $hit = $Preset.VArgs | Select-String -SimpleMatch "-pix_fmt" -Context 0,1 | Select-Object -First 1
        if ($hit -and $hit.Context -and $hit.Context.PostContext -and $hit.Context.PostContext.Count -ge 1) {
            $v = $hit.Context.PostContext[0]
            if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
        }
    } catch { }
    return "n/a"
}

function Resolve-ExeOrNull {
    param([Parameter(Mandatory)][string]$Name)
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) { return $cmd.Source }
    return $null
}

# Prefer system tools first
$ffmpegPath  = Resolve-ExeOrNull "ffmpeg"
$ffprobePath = Resolve-ExeOrNull "ffprobe"

# Upscayl: either user provided a path, or try PATH
$upscaylPathResolved = $null
if ($PSBoundParameters.ContainsKey("UpscaylPath") -and (Test-Path -LiteralPath $UpscaylPath)) {
    $upscaylPathResolved = (Resolve-Path -LiteralPath $UpscaylPath).Path
} else {
    $upscaylPathResolved = Resolve-ExeOrNull "upscayl-bin"
    if (-not $upscaylPathResolved) { $upscaylPathResolved = Resolve-ExeOrNull "upscayl-bin.exe" }
}

$needDownload = (-not $ffmpegPath) -or (-not $ffprobePath) -or (-not $upscaylPathResolved)

if ($needDownload) {
    if (-not $AutoDownloadTools) {
        throw ("Missing required tools. Install/provide them, or re-run with -AutoDownloadTools.`n" +
               "ffmpeg  : {0}`nffprobe : {1}`nupscayl : {2}" -f $ffmpegPath, $ffprobePath, $upscaylPathResolved)
    }

    # Only now do your Get-FromGitHubReleaseZip calls
    if ($AutoDownloadTools) {

        # ---------------------------
        # 1) FFmpeg + FFprobe from GyanD/codexffmpeg (*full_build.zip)
        # ---------------------------
        $ff = Get-FromGitHubReleaseZip `
        -Owner "GyanD" `
        -Repo "codexffmpeg" `
        -AssetNameRegex 'ffmpeg-.*-full_build(-shared)?\.zip$' `
        -InstallSubdir "ffmpeg" `
        -PrimaryExeName "ffmpeg.exe" `
        -SecondaryExeName "ffprobe.exe"

        # ---------------------------
        # 2) Upscayl (and whatever it bundles, potentially models) from upscayl/upscayl (*win.zip)
        # ---------------------------
        $up = Get-FromGitHubReleaseZip `
        -Owner "upscayl" `
        -Repo "upscayl" `
        -AssetNameRegex '^upscayl-\d+\.\d+\.\d+-win\.zip$' `
        -InstallSubdir "upscayl" `
        -PrimaryExeName "upscayl-bin.exe"

        # Override tool paths to downloaded versions
        $ffmpegPath  = $ff.PrimaryExe
        $ffprobePath = $ff.SecondaryExe
        $UpscaylPath = $up.PrimaryExe

    } else {
        Write-Host "[Tools] AutoDownloadTools not set; skipping downloads."
    }

    # Add downloaded tool dirs to PATH for this process
    $env:PATH = ((Split-Path $ffmpegPath -Parent) + ";" + (Split-Path $upscaylPathResolved -Parent) + ";" + $env:PATH)

    Write-Host "[Tools] Using downloaded tools:"
    Write-Host "  ffmpeg : $ffmpegPath ($($ff.Source))"
    Write-Host "  ffprobe: $ffprobePath ($($ff.Source))"
    Write-Host "  upscayl: $upscaylPathResolved ($($up.Source))"
} else {
    Write-Host "[Tools] Using system tools from PATH:"
    Write-Host "  ffmpeg : $ffmpegPath"
    Write-Host "  ffprobe: $ffprobePath"
    Write-Host "  upscayl: $upscaylPathResolved"
    
    # IMPORTANT: set the canonical vars the rest of the script uses
    $script:FfmpegExe  = $ffmpegPath
    $script:FfprobeExe = $ffprobePath
    $script:UpscaylExe = $upscaylPathResolved
}

# Now set the script variables you actually use
$Upscayl = $upscaylPathResolved


function Get-GitHubReleaseAsset {
    param(
        [Parameter(Mandatory)][string]$Owner,
        [Parameter(Mandatory)][string]$Repo,
        [Parameter(Mandatory)][string]$NameRegex,   # e.g. 'full_build\.zip$'
        [ValidateSet("latest","all")][string]$Mode = "latest"
    )

    $api =
        if ($Mode -eq "latest") {
            "https://api.github.com/repos/$Owner/$Repo/releases/latest"
        } else {
            "https://api.github.com/repos/$Owner/$Repo/releases"
        }

    $rel = Invoke-GitHubApi -Url $api

    # If Mode=all, $rel is an array; pick first release that has a matching asset
    $releases = @()
    if ($Mode -eq "all") { $releases = @($rel) } else { $releases = @($rel) }

    foreach ($r in $releases) {
        $asset = @($r.assets) | Where-Object { $_.name -match $NameRegex } | Select-Object -First 1
        if ($asset) {
            return [pscustomobject]@{
                Tag          = $r.tag_name
                Name         = $asset.name
                DownloadUrl  = $asset.browser_download_url
                SizeBytes    = $asset.size
                UpdatedAt    = $asset.updated_at
            }
        }
    }

    throw "No asset matching regex '$NameRegex' found for $Owner/$Repo."
}

function Invoke-DownloadFile {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$OutFile
    )
    New-Item -ItemType Directory -Force (Split-Path $OutFile -Parent) | Out-Null

    $bits = Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue
    if ($bits) {
        Start-BitsTransfer -Source $Url -Destination $OutFile -ErrorAction Stop
    } else {
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
    }
}

function Expand-ZipFresh {
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [Parameter(Mandatory)][string]$DestDir
    )
    if (Test-Path -LiteralPath $DestDir) {
        Remove-Item -LiteralPath $DestDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Force $DestDir | Out-Null
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $DestDir -Force
}

function Find-FirstFile {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$FileName
    )
    Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ieq $FileName } |
        Select-Object -First 1 -ExpandProperty FullName
}

function Get-FromGitHubReleaseZip {
    param(
        [Parameter(Mandatory)][string]$Owner,
        [Parameter(Mandatory)][string]$Repo,
        [Parameter(Mandatory)][string]$AssetNameRegex,
        [Parameter(Mandatory)][string]$InstallSubdir,
        [Parameter(Mandatory)][string]$PrimaryExeName,
        [string]$SecondaryExeName = ""
    )

    $installDir = Join-Path $ToolsRoot $InstallSubdir
    
    # Already installed?
    $installedPrimary = Find-FirstFile -Root $installDir -FileName $PrimaryExeName
    if ($installedPrimary) {
        $installedSecondary = $null
        if ($SecondaryExeName) {
            $installedSecondary = Find-FirstFile -Root $installDir -FileName $SecondaryExeName
        }

        return [pscustomobject]@{
            InstallDir   = $installDir
            PrimaryExe   = $installedPrimary
            SecondaryExe = $installedSecondary
            Source       = "cache"
        }
    }

    $asset = Get-GitHubReleaseAsset -Owner $Owner -Repo $Repo -NameRegex $AssetNameRegex
    Write-Host "[Tools] Downloading $Owner/$Repo $($asset.Tag) asset: $($asset.Name)"

    $cacheDir = Join-Path $ToolsRoot "_cache"
    New-Item -ItemType Directory -Force $cacheDir | Out-Null
    $zipPath  = Join-Path $cacheDir $asset.Name

    Invoke-DownloadFile -Url $asset.DownloadUrl -OutFile $zipPath

    # Extract to temp then copy the folder that contains the exe(s)
    $tmp = Join-Path $cacheDir ("extract_{0}_{1}" -f $InstallSubdir, ([guid]::NewGuid().ToString("N")))
    Expand-ZipFresh -ZipPath $zipPath -DestDir $tmp

    $foundPrimary = Find-FirstFile -Root $tmp -FileName $PrimaryExeName
    if (-not $foundPrimary) {
        throw "Downloaded $asset.Name but couldn't find $PrimaryExeName inside it."
    }

    $srcDir = Split-Path $foundPrimary -Parent

    # If we also need a secondary exe (ffprobe), ensure it’s in same folder or nearby
    if ($SecondaryExeName) {
        $foundSecondary = Find-FirstFile -Root $tmp -FileName $SecondaryExeName
        if (-not $foundSecondary) {
            throw "Downloaded $asset.Name but couldn't find $SecondaryExeName inside it."
        }
        $srcDir2 = Split-Path $foundSecondary -Parent
        if ($srcDir2 -ne $srcDir) {
            # If they are in different folders, copy both folder trees into installDir (simplest safe approach)
            # We'll just copy from the common root to preserve needed DLLs/resources.
            $srcDir = $tmp
        }
    }

    # Install: copy the whole folder contents (keeps DLLs/models/etc)
    if (Test-Path -LiteralPath $installDir) {
        Remove-Item -LiteralPath $installDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Force $installDir | Out-Null
    Get-ChildItem -LiteralPath $srcDir -Force -ErrorAction Stop |
    Copy-Item -Destination $installDir -Recurse -Force -ErrorAction Stop


    # Re-locate exe(s) inside install dir (in case they are in nested bin\)
    $installedPrimary = Find-FirstFile -Root $installDir -FileName $PrimaryExeName
    if (-not $installedPrimary) { throw "Install succeeded but $PrimaryExeName not found under $installDir" }

    $installedSecondary = $null
    if ($SecondaryExeName) {
        $installedSecondary = Find-FirstFile -Root $installDir -FileName $SecondaryExeName
        if (-not $installedSecondary) { throw "Install succeeded but $SecondaryExeName not found under $installDir" }
    }

    # Cleanup temp extract
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue

    return [pscustomobject]@{
        InstallDir   = $installDir
        PrimaryExe   = $installedPrimary
        SecondaryExe = $installedSecondary
        Source      = "${Owner}/${Repo}:$($asset.Tag)"
    }
}

function Get-TargetDimsFromPreset {
    param([Parameter(Mandatory)][pscustomobject]$Preset)

    if ([string]::IsNullOrWhiteSpace($Preset.Scale)) {
        return $null
    }

    $p = $Preset.Scale -split ":"
    if ($p.Count -ne 2) { throw "Preset.Scale must be 'W:H' but was: $($Preset.Scale)" }

    return [pscustomobject]@{
        W = [int]$p[0]
        H = [int]$p[1]
    }
}

# ---------------------------
# Canonical tool exe paths (downloaded)
# ---------------------------

$script:FfmpegExe  = $ff.PrimaryExe
$script:FfprobeExe = $ff.SecondaryExe
$script:UpscaylExe = $up.PrimaryExe

# NOTE: validate *resolved paths*, not $script:* (those are set later)
if ([string]::IsNullOrWhiteSpace($ffmpegPath)  -or -not (Test-Path -LiteralPath $ffmpegPath))  { throw "ffmpeg not found: '$ffmpegPath'" }
if ([string]::IsNullOrWhiteSpace($ffprobePath) -or -not (Test-Path -LiteralPath $ffprobePath)) { throw "ffprobe not found: '$ffprobePath'" }
if ([string]::IsNullOrWhiteSpace($upscaylPathResolved) -or -not (Test-Path -LiteralPath $upscaylPathResolved)) { throw "upscayl not found: '$upscaylPathResolved'" }

$ffDir = Split-Path -Parent $ffmpegPath
$upDir = Split-Path -Parent $upscaylPathResolved
$env:PATH = "$ffDir;$upDir;$env:PATH"

# If user explicitly provided -UpscaylPath, honor it (but fall back to downloaded if not found)
if ($PSBoundParameters.ContainsKey('UpscaylPath')) {
    try {
        $resolvedUserUpscayl = Resolve-FullPath $script:UpscaylExe
        if (Test-Path -LiteralPath $resolvedUserUpscayl) {
            $script:UpscaylExe = $resolvedUserUpscayl
        } else {
            Write-Warning "User-specified UpscaylPath not found; using downloaded upscayl: $script:UpscaylExe"
        }
    } catch {
        Write-Warning "Could not resolve user UpscaylPath; using downloaded upscayl: $script:UpscaylExe"
    }
}

$ffDir = if ($script:FfmpegExe)  { Split-Path -Parent $script:FfmpegExe }  else { $null }
$upDir = if ($script:UpscaylExe) { Split-Path -Parent $script:UpscaylExe } else { $null }

if ($ffDir -and $upDir) {
    $env:PATH = "$ffDir;$upDir;$env:PATH"
}


# ---- Canonicalize tooling vars (single source of truth) ----
$script:FfmpegExe  = $ffmpegPath
$script:FfprobeExe = $ffprobePath
$script:UpscaylExe = $upscaylPathResolved

if ([string]::IsNullOrWhiteSpace($script:FfmpegExe)  -or -not (Test-Path -LiteralPath $script:FfmpegExe))  {
    throw "ffmpeg not found/resolved: '$script:FfmpegExe'"
}
if ([string]::IsNullOrWhiteSpace($script:FfprobeExe) -or -not (Test-Path -LiteralPath $script:FfprobeExe)) {
    throw "ffprobe not found/resolved: '$script:FfprobeExe'"
}
if ([string]::IsNullOrWhiteSpace($script:UpscaylExe) -or -not (Test-Path -LiteralPath $script:UpscaylExe)) {
    throw "upscayl not found/resolved: '$script:UpscaylExe'"
}

# Add tool dirs to PATH (nice-to-have)
$ffDir = Split-Path -Parent $script:FfmpegExe
$upDir = Split-Path -Parent $script:UpscaylExe
$env:PATH = "$ffDir;$upDir;$env:PATH"

# InputDir: default "source" already; resolve relative
$InputDirFull = Join-Path $BaseDir $InputDir
if (-not (Test-Path $InputDirFull)) { throw "Source directory not found: $InputDirFull" }
$InputDirFull = (Resolve-Path $InputDirFull).Path

# ---- Interactive fallback ONLY if not provided explicitly ----

if (-not $PSBoundParameters.ContainsKey("InputDir")) {
    $raw = Read-Host "Enter source directory (default: source)"
    if (-not [string]::IsNullOrWhiteSpace($raw)) {
        $InputDir = $raw
        $InputDirFull = (Resolve-Path (Join-Path $BaseDir $InputDir)).Path
    }
}

$GpuTag = if ($GpuIndex -ge 0) { "gpu$GpuIndex" } else { "gpuNA" }

# ---------------- FUNCTIONS ----------------
function Invoke-Ffmpeg  { & $script:FfmpegExe  @args }
function Invoke-Ffprobe { & $script:FfprobeExe @args }
function Invoke-Upscayl { & $script:UpscaylExe @args }
function Get-UpscaylHelpText {
    param([Parameter(Mandatory)][string]$script:UpscaylExe)

    # cmd.exe captures stderr+stdout cleanly without PS "NativeCommandError" noise
    $escaped = $script:UpscaylExe.Replace('"','""')
    return (cmd /c """$escaped"" -h 2>&1") | Out-String
}

function Get-SafeName([string]$Name) {
  $n = $Name -replace '[\p{Cc}]',''           # drop control chars
  $n = $n -replace '[<>:"/\\|?*]','_'         # Windows-illegal in paths
  $n = $n.Trim().TrimEnd('.')                 # avoid trailing dot
  if ([string]::IsNullOrWhiteSpace($n)) { $n = "video" }
  return $n
}

function Get-PixFmtFromVArgs {
    param([object[]]$VArgs)

    if (-not $VArgs) { return "n/a" }

    for ($i = 0; $i -lt $VArgs.Count; $i++) {
        if ($VArgs[$i] -eq "-pix_fmt") {
            $j = $i + 1
            if ($j -lt $VArgs.Count -and $VArgs[$j]) { return [string]$VArgs[$j] }
        }
    }
    return "n/a"
}

function Get-CounterSamples {
    param([Parameter(Mandatory)][string]$Path)
    try {
        return (Get-Counter -Counter $Path -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop).CounterSamples
    } catch {
        return @()  # counter missing or inaccessible
    }
}
function Get-GpuDedicatedVramPercent {
    <#
      Returns a hashtable: gpuIndex -> dedicated VRAM percent used (0-100)
      Uses Windows perf counters:
        \GPU Adapter Memory(*)\Dedicated Usage
        \GPU Adapter Memory(*)\Dedicated Limit   (may not exist!)
    #>

    $useS = Get-CounterSamples '\GPU Adapter Memory(*)\Dedicated Usage'
    $limS = Get-CounterSamples '\GPU Adapter Memory(*)\Dedicated Limit'  # may be empty on some systems

    if (-not $useS -or $useS.Count -eq 0) { return @{} }

    # Map instance -> max usage/limit
    $useMap = @{}
    foreach ($s in $useS) { $useMap[$s.InstanceName] = [double]$s.CookedValue }

    $limMap = @{}
    foreach ($s in $limS) { $limMap[$s.InstanceName] = [double]$s.CookedValue }

    # Aggregate by GPU index (gpu_#) across instances
    $agg = @{}
    foreach ($inst in $useMap.Keys) {
        if ($inst -match 'gpu_(\d+)') {
            $idx = [int]$Matches[1]
            if (-not $agg.ContainsKey($idx)) { $agg[$idx] = [pscustomobject]@{ Use=0.0; Lim=0.0 } }

            $agg[$idx].Use = [math]::Max($agg[$idx].Use, $useMap[$inst])
            if ($limMap.ContainsKey($inst)) {
                $agg[$idx].Lim = [math]::Max($agg[$idx].Lim, $limMap[$inst])
            }
        }
    }

    # Convert to percent (only where limit exists)
    $out = @{}
    foreach ($k in $agg.Keys) {
        $useB = $agg[$k].Use
        $limB = $agg[$k].Lim
        if ($limB -gt 0) {
            $out[$k] = [math]::Round(($useB / $limB) * 100, 1)
        }
    }

    return $out
}

function Get-LockedGpus {
    $locked = @()
    Get-ChildItem -Directory "_ffv1_work_gpu*" -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.Name -match 'gpu(\d+)') {
            # Non-empty directory = active or incomplete job
            $hasFiles = Get-ChildItem $_.FullName -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($hasFiles) {
                $locked += [int]$Matches[1]
            }
        }
    }
    return $locked
}

function Get-GpuUtilizationByLuid {
    # Returns: LUID -> @{ Util3D = x; UtilCompute = y; UtilMax = z }
    try {
        $c = Get-CounterSamples '\GPU Engine(*)\Utilization Percentage' -SampleInterval 1 -MaxSamples 1
    } catch {
        return @{}
    }

    $out = @{}

    foreach ($s in $c.CounterSamples) {
        $inst = $s.InstanceName

        # Extract LUID
        if ($inst -notmatch '(luid_0x[0-9a-fA-F]+_0x[0-9a-fA-F]+)') { continue }
        $luid = $Matches[1]

        # Identify engine type
        $is3D      = ($inst -like '*engtype_3D*')
        $isCompute = ($inst -like '*engtype_Compute*')

        if (-not ($is3D -or $isCompute)) { continue }

        if (-not $out.ContainsKey($luid)) {
            $out[$luid] = [pscustomobject]@{ Util3D = 0.0; UtilCompute = 0.0; UtilMax = 0.0 }
        }

        $val = [double]$s.CookedValue

        if ($is3D) {
            # Use MAX (summing can exceed 100 with multiple engines)
            $out[$luid].Util3D = [math]::Max($out[$luid].Util3D, $val)
        }
        if ($isCompute) {
            $out[$luid].UtilCompute = [math]::Max($out[$luid].UtilCompute, $val)
        }

        $out[$luid].UtilMax = [math]::Max($out[$luid].UtilMax, $val)
    }

    return $out
}

function Get-DiscreteVideoControllers {
    $vc = Get-CimInstance Win32_VideoController | Select-Object Name, AdapterCompatibility, PNPDeviceID, AdapterRAM

    # Filter out obvious non-GPU / useless adapters
    $vc = $vc | Where-Object {
        $_.Name -and
        $_.Name -notmatch 'Basic Display|Microsoft Remote|DisplayLink|Virtual|VMware|Hyper-V|Citrix|BCM|Broadcom' -and
        $_.AdapterCompatibility -notmatch 'Microsoft'
    }

    # Prefer AMD/NVIDIA and larger VRAM
    $vc | Sort-Object @{
        Expression = {
            $score = 0
            if ($_.AdapterCompatibility -match 'NVIDIA|AMD|Advanced Micro Devices') { $score += 100 }
            if ($_.Name -match 'NVIDIA|GeForce|RTX|Quadro') { $score += 50 }
            if ($_.Name -match 'AMD|Radeon') { $score += 50 }
            $ramGB = 0
            if ($_.AdapterRAM) { $ramGB = [math]::Round($_.AdapterRAM / 1GB, 0) }
            $score += [int]$ramGB
            $score
        }
        Descending = $true
    }
}

function Get-ActiveGpuLuids {
    param(
        [double]$MinVramGB = 2.0,
        [double]$MinUtilPercent = 5.0
    )

    $util = Get-GpuUtilizationByLuid
    $mem  = Get-GpuDedicatedUsageByLuid

    $luids = New-Object System.Collections.Generic.List[string]

    foreach ($luid in (($util.Keys + $mem.Keys) | Sort-Object -Unique)) {
        $useGB = if ($mem.ContainsKey($luid)) { $mem[$luid].UseGB } else { 0.0 }
        $umax  = if ($util.ContainsKey($luid)) { $util[$luid].UtilMax } else { 0.0 }

        if ($useGB -ge $MinVramGB -or $umax -ge $MinUtilPercent) {
            $luids.Add($luid)
        }
    }
    return $luids
}


function Get-GpuDedicatedUsageByLuid {
    $useS = Get-CounterSamples '\GPU Adapter Memory(*)\Dedicated Usage'
    $limS = Get-CounterSamples '\GPU Adapter Memory(*)\Dedicated Limit'  # may be empty

    $limMap = @{}
    foreach ($s in $limS) {
        if ($s.InstanceName -match '(luid_0x[0-9a-fA-F]+_0x[0-9a-fA-F]+)') {
            $limMap[$Matches[1]] = [double]$s.CookedValue
        }
    }

    $out = @{}
    foreach ($s in $useS) {
        if ($s.InstanceName -match '(luid_0x[0-9a-fA-F]+_0x[0-9a-fA-F]+)') {
            $luid = $Matches[1]
            $useB = [double]$s.CookedValue
            $limB = if ($limMap.ContainsKey($luid)) { [double]$limMap[$luid] } else { 0.0 }

            $out[$luid] = [pscustomobject]@{
                UseGB = [math]::Round($useB / 1GB, 2)
                LimGB = [math]::Round($limB / 1GB, 2)
                Pct   = if ($limB -gt 0) { [math]::Round(($useB / $limB) * 100, 1) } else { $null }
            }
        }
    }
    return $out
}


function Get-Gpu3DUtilization {
    try {
        $c = Get-CounterSamples '\GPU Engine(*)\Utilization Percentage' -SampleInterval 1 -MaxSamples 1
    } catch {
        return @{}
    }

    $util = @{}

    $c.CounterSamples |
        Where-Object { $_.InstanceName -like '*engtype_3D*' } |
        ForEach-Object {
            if ($_.InstanceName -match 'gpu_(\d+)_') {
                $idx = [int]$Matches[1]
                if (-not $util.ContainsKey($idx)) { $util[$idx] = 0 }
                $util[$idx] = [math]::Max($util[$idx], [double]$_.CookedValue)
            }
        }

    # Round for readability
    foreach ($k in $util.Keys) {
        $util[$k] = [math]::Round($util[$k], 1)
    }

    return $util
}

function Test-AnyGpuBusyOrLocked {
    param(
        [int]$BusyThresholdPercent = 50,
        [int]$BusyVramGB = 18,
        [double]$MinActiveVramGB = 2.0,
        [double]$MinActiveUtilPercent = 5.0
    )

    $locked = Get-LockedGpus
    if ($locked.Count -gt 0) {
        Write-Warning ("Locked GPUs (work dirs present): {0}" -f ($locked -join ", "))
    }

    $util = Get-GpuUtilizationByLuid
    $mem  = Get-GpuDedicatedUsageByLuid

    # Only consider "real" adapters (filters out BCM / 0GB junk)
    $activeLuids = Get-ActiveGpuLuids -MinVramGB $MinActiveVramGB -MinUtilPercent $MinActiveUtilPercent
    if (-not $activeLuids -or $activeLuids.Count -eq 0) {
        Write-Host "`n[GPU Busy Check] (No active GPU LUIDs detected from counters)"
        return
    }

    Write-Host "`n[GPU Busy Check | by LUID]"

    $warn = @()

    foreach ($luid in ($activeLuids | Sort-Object)) {
        $u3d  = if ($util.ContainsKey($luid)) { [math]::Round($util[$luid].Util3D, 2) } else { 0.0 }
        $ucmp = if ($util.ContainsKey($luid)) { [math]::Round($util[$luid].UtilCompute, 2) } else { 0.0 }
        $umax = if ($util.ContainsKey($luid)) { [math]::Round($util[$luid].UtilMax, 2) } else { 0.0 }

        $useGB = if ($mem.ContainsKey($luid)) { $mem[$luid].UseGB } else { 0.0 }
        $limGB = if ($mem.ContainsKey($luid)) { $mem[$luid].LimGB } else { 0.0 }
        $pct   = if ($mem.ContainsKey($luid)) { $mem[$luid].Pct } else { $null }

        $pctText = if ($null -ne $pct) { "$pct%" } else { "n/a" }

        Write-Host ("  {0,-30}  3D={1,7}%  Compute={2,7}%  Max={3,7}%  VRAM={4,6}GB / {5,6}GB ({6})" -f `
            $luid, $u3d, $ucmp, $umax, $useGB, $limGB, $pctText)

        if ($umax -ge $BusyThresholdPercent) {
            $warn += "LUID $luid utilization is high (MaxEng=$umax%)"
        }

        if ($null -ne $pct) {
            if ($pct -ge $BusyThresholdPercent) { $warn += "LUID $luid VRAM is high ($pct%)" }
        } else {
            if ($useGB -ge $BusyVramGB) { $warn += "LUID $luid VRAM is high by usage ($useGB GB ≥ $BusyVramGB GB) (limit counter unavailable)" }
        }
    }

    if ($warn.Count -gt 0) {
        Write-Warning "One or more GPUs look busy:"
        $warn | ForEach-Object { Write-Warning "  - $_" }

        $resp = Read-Host "Continue anyway? (Y/N)"
        if ($resp.Trim().ToUpper() -ne "Y") { throw "Aborted due to GPU busy warning" }
    }
}


function Get-GpuMemoryByLuid {
    $useS = Get-CounterSamples '\GPU Adapter Memory(*)\Dedicated Usage'
    $limS = Get-CounterSamples '\GPU Adapter Memory(*)\Dedicated Limit'  # may be empty

    if (-not $useS -or $useS.Count -eq 0) { return @() }

    $limMap = @{}
    foreach ($s in $limS) {
        if ($s.InstanceName -match '(luid_0x[0-9a-fA-F]+_0x[0-9a-fA-F]+)') {
            $limMap[$Matches[1]] = [double]$s.CookedValue
        }
    }

    $out = @()
    foreach ($s in $useS) {
        if ($s.InstanceName -match '(luid_0x[0-9a-fA-F]+_0x[0-9a-fA-F]+)') {
            $luid = $Matches[1]
            $useB = [double]$s.CookedValue
            $limB = if ($limMap.ContainsKey($luid)) { [double]$limMap[$luid] } else { 0.0 }

            $out += [pscustomobject]@{
                Luid = $luid
                DedicatedUsageBytes = $useB
                DedicatedLimitBytes = $limB
            }
        }
    }

    return $out
}

function Get-GpuEngine3DUtilByLuid {
    # Returns objects keyed by LUID with summed 3D utilization (0-100)
    # Counter instance name looks like:
    # "pid_0x0_luid_0x00000000_0x0000..._phys_0_engtype_3D"
    try {
        $c = Get-CounterSamples '\GPU Engine(*)\Utilization Percentage' -SampleInterval 1 -MaxSamples 1
    } catch {
        return @()
    }

    $samples = $c.CounterSamples | Where-Object {
        $_.InstanceName -like '*engtype_3D*'
    }

    $groups = $samples | Group-Object -Property {
        if ($_.InstanceName -match '(luid_0x[0-9a-fA-F]+_0x[0-9a-fA-F]+)') { $Matches[1] } else { 'unknown' }
    }

    foreach ($g in $groups) {
        [pscustomobject]@{
            Luid = $g.Name
            Util3D = [math]::Round( ($g.Group | Measure-Object -Property CookedValue -Sum).Sum, 2 )
        }
    }
}

function Show-GpuEngineSnapshot {
    $util = Get-GpuEngine3DUtilByLuid
    $mem  = Get-GpuMemoryByLuid

    Write-Host "`n[GPU Engine 3D Utilization by LUID]"
    if (-not $util -or $util.Count -eq 0) {
        Write-Host "  (No GPU Engine counters available)"
    } else {
        $util | Sort-Object Util3D -Descending | ForEach-Object {
            Write-Host ("  {0,-30}  {1,8} %" -f $_.Luid, $_.Util3D)
        }
    }

    Write-Host "`n[GPU Dedicated Memory by LUID]"
    if (-not $mem -or $mem.Count -eq 0) {
        Write-Host "  (No GPU Adapter Memory counters available)"
    } else {
        $mem | Sort-Object DedicatedLimitBytes -Descending | ForEach-Object {
            $uGB = [math]::Round($_.DedicatedUsageBytes / 1GB, 2)
            $lGB = if ($_.DedicatedLimitBytes -gt 0) { [math]::Round($_.DedicatedLimitBytes / 1GB, 2) } else { $null }
            Write-Host ("  {0,-30}  {1,7} GB / {2,7} GB" -f $_.Luid, $uGB, $lGB)
        }
    }
}
function Get-UpscaylGpuList {
    param([Parameter(Mandatory)][string]$script:UpscaylExe)

    $out = & $script:UpscaylExe -h 2>&1 | Out-String
    if (-not $out) { $out = & $script:UpscaylExe 2>&1 | Out-String }

    $gpus = @()

    foreach ($line in ($out -split "`r?`n")) {
        $t = $line.Trim()

        # Accept formats like:
        # [0 AMD Radeon ...]
        # 0: AMD Radeon ...
        # GPU 0 - AMD Radeon ...
        if ($t -match '^\[(\d+)\s+(.+?)\]') {
            $gpus += [pscustomobject]@{ Index=[int]$Matches[1]; Name=$Matches[2].Trim(); Raw=$t }
            continue
        }
        if ($t -match '^(?:GPU\s*)?(\d+)\s*[:\-]\s*(.+)$') {
            $gpus += [pscustomobject]@{ Index=[int]$Matches[1]; Name=$Matches[2].Trim(); Raw=$t }
            continue
        }
    }

    return $gpus
}


function Select-UpscaylGpuIndexAuto {
    param(
        [Parameter(Mandatory)][string]$script:UpscaylExe,
        [switch]$PreferDiscrete
    )

    $up = Get-UpscaylGpuList -UpscaylPath $script:UpscaylExe
    if (-not $up -or $up.Count -eq 0) {
        Write-Warning "Couldn't parse Upscayl GPU list. Falling back to GPU 0."
        return 0
    }

    $vc = Get-DiscreteVideoControllers

    # Util + VRAM by LUID (Windows counters)
    #$util = Get-GpuEngine3DUtilByLuid
    #$mem  = Get-GpuMemoryByLuid

    # Build a "discrete-ish" guess list from Windows adapters
    # (AMD/NVIDIA typically discrete; Intel typically integrated)
    $discreteNames = @()
    if ($PreferDiscrete) {
        $discreteNames = $vc | Where-Object {
            $_.AdapterCompatibility -match 'AMD|Advanced Micro Devices|NVIDIA' -or $_.Name -match 'AMD|Radeon|NVIDIA|GeForce|RTX|GTX'
        } | ForEach-Object { $_.Name }
    }

    # If we can’t map LUID -> Upscayl index perfectly, we do best-effort:
    # - Prefer Upscayl device names that match a discrete adapter name
    # - Use engine utilization totals as a tie-breaker if available (global low utilization)
    # - Prefer higher Dedicated Limit (discrete tends to have large dedicated limit)
    #
    # We can’t directly map Upscayl index to LUID without extra DXGI plumbing,
    # so we use a pragmatic approach:
    #   choose the Upscayl GPU that matches a discrete name first.
    $candidates = $up

    if ($PreferDiscrete -and $discreteNames.Count -gt 0) {
        $disc = @()
        foreach ($g in $up) {
            foreach ($n in $discreteNames) {
                if ($g.Name -and $n -and ($g.Name -like "*$($n.Split(' ')[0])*" -or $n -like "*$($g.Name.Split(' ')[0])*" -or $g.Name -match 'Radeon|NVIDIA|GeForce|RTX|GTX')) {
                    $disc += $g
                    break
                }
            }
        }
        if ($disc.Count -gt 0) { $candidates = $disc }
    }

    # If we have memory counters, prefer the adapter with the largest dedicated limit
    # (this usually corresponds to the discrete card). We can only use this as a heuristic:
    $best = $candidates | Select-Object -First 1

    # If multiple candidates, pick the one that "looks" best by name (XTX/XT/RTX etc.)
    if ($candidates.Count -gt 1) {
        $best = $candidates |
          Sort-Object -Property @{
              Expression = {
                  # crude scoring by keywords
                  $s = 0
                  if ($_.Name -match 'XTX|XT|RTX|4090|4080|4070|7900|7800|7700|6950|6900') { $s += 10 }
                  if ($_.Name -match 'Radeon|NVIDIA|GeForce') { $s += 5 }
                  $s
              }
              Descending = $true
          } | Select-Object -First 1
    }

    return $best.Index
}

# ---------------- GPU SELECTION (after functions exist) ----------------

# Only pick GPU if we actually need it for work.
# For pure -DryRun, we can skip GPU logic entirely.
if (-not $DryRun) {

    if ($GpuIndex -lt 0) {
        $gpuInput = Read-Host "GPU index for Upscayl (Enter = auto)"
        if ([string]::IsNullOrWhiteSpace($gpuInput)) {
            $GpuIndex = -1   # auto
        } else {
            if ($gpuInput -notmatch '^\d+$') { throw "GPU index must be numeric" }
            $GpuIndex = [int]$gpuInput
        }

    }

    # -1 means "Upscayl auto" and is allowed.

    Test-AnyGpuBusyOrLocked -BusyThresholdPercent $BusyThresholdPercent
}

# ---------------- END GPU SELECTION ----------------

$GpuTag = if ($GpuIndex -ge 0) { "gpu$GpuIndex" } else { "gpuAuto" }

# Compute WorkDir/OutputDir if not provided
if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $WorkDir = Join-Path $BaseDir "_ffv1_work_$GpuTag"
} else {
    $WorkDir = Join-Path $BaseDir $WorkDir
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    switch ($OutputPreset) {
        "h264_4k_mp4"     { $OutputDir = Join-Path $BaseDir "h264_4k_$GpuTag" }
        "hevc_4k_mp4"     { $OutputDir = Join-Path $BaseDir "hevc_4k_$GpuTag" }
        "prores_4k_mov"   { $OutputDir = Join-Path $BaseDir "prores_4k_$GpuTag" }
        "av1_4k_mkv"      { $OutputDir = Join-Path $BaseDir "av1_4k_$GpuTag" }
        "av1_4k_mp4"      { $OutputDir = Join-Path $BaseDir "av1_4k_$GpuTag" }
        "ffv1_1080p_mkv" { $OutputDir = Join-Path $BaseDir "ffv1_1080p_$GpuTag" }
        "ffv1_4k_mkv"    { $OutputDir = Join-Path $BaseDir "ffv1_4k_$GpuTag" }
        "av1_1080p_mkv"  { $OutputDir = Join-Path $BaseDir "av1_1080p_$GpuTag" }
        "av1_1080p_mp4"  { $OutputDir = Join-Path $BaseDir "av1_1080p_$GpuTag" }
        default          { $OutputDir = Join-Path $BaseDir "output_$OutputPreset`_$GpuTag" }
    }
} else {
    $OutputDir = Join-Path $BaseDir $OutputDir
}

New-Item -ItemType Directory -Force $WorkDir   | Out-Null
New-Item -ItemType Directory -Force $OutputDir | Out-Null

# Replace your old $InputDir usage with $InputDirFull going forward
$InputDir = $InputDirFull
$Upscayl = $script:UpscaylExe

function Test-OutputValid {
    param(
        [Parameter(Mandatory)][string]$file,
        [switch]$RequireAudio
    )
    if (-not (Test-Path -LiteralPath $file)) { return $false }

    # ask ffprobe for codec_type for ALL streams
    $types = & $script:FfprobeExe -v error `
        -show_entries stream=codec_type `
        -of csv=p=0 `
        "file:$file" 2>$null

    if (-not $types) { return $false }

    $hasVideo = ($types -match 'video')
    $hasAudio = ($types -match 'audio')

    if (-not $hasVideo) { return $false }
    if ($RequireAudio -and (-not $hasAudio)) { return $false }

    return $true
}


function Get-OutputPresetConfig {
    param(
        [Parameter(Mandatory)][string]$PresetName
    )

    switch ($PresetName) {

        "ffv1_1080p_mkv" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mkv"
                VCodec    = "ffv1"
                VArgs     = @("-level","3","-g","1","-pix_fmt","yuv420p")
                ACodec    = "flac"
                AArgs     = @()
                Scale     = "1920:1080"
                ExtraMaps = @() # reserved
            }
        }

        "ffv1_4k_mkv" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mkv"
                VCodec    = "ffv1"
                VArgs     = @("-level","3","-g","1","-pix_fmt","yuv420p")
                ACodec    = "flac"
                AArgs     = @()
                Scale     = "3840:2160"
                ExtraMaps = @()
            }
        }

        "h264_1080p_mp4" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mp4"
                VCodec    = "libx264"
                VArgs     = @("-pix_fmt","yuv420p","-preset","slow","-crf","18","-profile:v","high","-movflags","+faststart")
                ACodec    = "aac"
                AArgs     = @("-b:a","192k")
                Scale     = "1920:1080"
                ExtraMaps = @()
            }
        }

        "hevc_1080p_mp4" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mp4"
                VCodec    = "libx265"
                VArgs     = @("-pix_fmt","yuv420p","-preset","slow","-crf","20","-tag:v","hvc1","-movflags","+faststart")
                ACodec    = "aac"
                AArgs     = @("-b:a","192k")
                Scale     = "1920:1080"
                ExtraMaps = @()
            }
        }

        "prores_1080p_mov" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mov"
                VCodec    = "prores_ks"
                VArgs     = @("-profile:v","3","-pix_fmt","yuv422p10le") # ProRes 422 HQ
                ACodec    = "pcm_s16le"
                AArgs     = @()
                Scale     = "1920:1080"
                ExtraMaps = @()
            }
        }

        "av1_1080p_mkv" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mkv"
                VCodec    = "libsvtav1"
                # Good quality-per-size starter settings:
                # CRF ~28 is a common balance; preset 6 is a reasonable speed/efficiency compromise.
                VArgs     = @("-pix_fmt","yuv420p","-preset","6","-crf","28")
                ACodec    = "libopus"
                AArgs     = @("-b:a","128k")
                Scale     = "1920:1080"
                ExtraMaps = @()
            }
        }
        
        "av1_1080p_mp4" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mp4"
                VCodec    = "libsvtav1"
                # MP4 needs AV1 tag; +faststart is good for playback.
                VArgs     = @("-pix_fmt","yuv420p","-preset","6","-crf","28","-tag:v","av01","-movflags","+faststart")
                ACodec    = "aac"
                AArgs     = @("-b:a","160k")
                Scale     = "1920:1080"
                ExtraMaps = @()
            }
        }
        "h264_4k_mp4" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mp4"
                VCodec    = "libx264"
                # 4K delivery: raise CRF slightly vs 1080p to keep size sane
                VArgs     = @("-pix_fmt","yuv420p","-preset","slow","-crf","20","-profile:v","high","-movflags","+faststart")
                ACodec    = "aac"
                AArgs     = @("-b:a","192k")
                Scale     = "3840:2160"
                ExtraMaps = @()
            }
        }

        "hevc_4k_mp4" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mp4"
                VCodec    = "libx265"
                # 4K HEVC: good size/quality; keep hvc1 tag for Apple players
                VArgs     = @("-pix_fmt","yuv420p","-preset","slow","-crf","22","-tag:v","hvc1","-movflags","+faststart")
                ACodec    = "aac"
                AArgs     = @("-b:a","192k")
                Scale     = "3840:2160"
                ExtraMaps = @()
            }
        }

        "prores_4k_mov" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mov"
                VCodec    = "prores_ks"
                # ProRes 422 HQ (profile 3), 10-bit 4:2:2
                VArgs     = @("-profile:v","3","-pix_fmt","yuv422p10le")
                ACodec    = "pcm_s16le"
                AArgs     = @()
                Scale     = "3840:2160"
                ExtraMaps = @()
            }
        }

        "av1_4k_mkv" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mkv"
                VCodec    = "libsvtav1"
                # 4K AV1: a little lower quality setting (higher CRF) keeps size reasonable
                VArgs     = @("-pix_fmt","yuv420p","-preset","6","-crf","30")
                ACodec    = "libopus"
                AArgs     = @("-b:a","160k")
                Scale     = "3840:2160"
                ExtraMaps = @()
            }
        }

        "av1_4k_mp4" {
            return [pscustomobject]@{
                Name      = $PresetName
                Ext       = "mp4"
                VCodec    = "libsvtav1"
                VArgs     = @("-pix_fmt","yuv420p","-preset","6","-crf","30","-tag:v","av01","-movflags","+faststart")
                ACodec    = "aac"
                AArgs     = @("-b:a","160k")
                Scale     = "3840:2160"
                ExtraMaps = @()
            }
        }

        default {
            throw "Unknown OutputPreset: $PresetName"
        }
    }
}
function Get-OutputPathForPreset {
    param(
        [Parameter(Mandatory)][string]$OutputDir,
        [Parameter(Mandatory)][string]$BaseName,
        [Parameter(Mandatory)][string]$GpuTag,
        [Parameter(Mandatory)][pscustomobject]$Preset
    )

    $suffix = $Preset.Name
    return (Join-Path $OutputDir ("{0}_{1}_{2}.{3}" -f $BaseName, $suffix, $GpuTag, $Preset.Ext))
}

function Get-VideoInfo {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }

    $json = & $script:FfprobeExe -v error `
        -print_format json `
        -show_format `
        -show_streams `
        "file:$Path" | Out-String

    if (-not $json) { throw "ffprobe returned no data for: $Path" }

    $o = $json | ConvertFrom-Json

    $v = $o.streams | Where-Object { $_.codec_type -eq "video" } | Select-Object -First 1
    $a = $o.streams | Where-Object { $_.codec_type -eq "audio" } | Select-Object -First 1

    $fps = 30.0
    if ($v.r_frame_rate) {
        if ($v.r_frame_rate -match '^(\d+)/(\d+)$') {
            $n = [double]$Matches[1]; $d = [double]$Matches[2]
            if ($d -ne 0) { $fps = $n / $d }
        } else {
            [double]::TryParse($v.r_frame_rate, [ref]$fps) | Out-Null
        }
    }

    [pscustomobject]@{
        Path         = $Path
        Container    = $o.format.format_name
        DurationSec  = [double]$o.format.duration
        BitrateKbps  = [int]([double]$o.format.bit_rate / 1000)

        VCodec       = $v.codec_name
        Width        = [int]$v.width
        Height       = [int]$v.height
        PixFmt       = $v.pix_fmt
        ColorSpace   = $v.color_space
        ColorPrim    = $v.color_primaries
        ColorTrc     = $v.color_transfer
        Fps          = [math]::Round($fps, 6)

        ACodec       = $a.codec_name
        ASampleRate  = $a.sample_rate
        AChannels    = $a.channels
    }
}

function Show-VideoInfo {
    param([Parameter(Mandatory)]$Info)

    Write-Host ""
    Write-Host "[Detected Source]"
    Write-Host ("  Container : {0}" -f $Info.Container)
    Write-Host ("  Video     : {0}  {1}x{2}  {3}  fps={4}" -f $Info.VCodec, $Info.Width, $Info.Height, $Info.PixFmt, $Info.Fps)
    Write-Host ("  Color     : space={0} prim={1} trc={2}" -f $Info.ColorSpace, $Info.ColorPrim, $Info.ColorTrc)
    Write-Host ("  Audio     : {0}  {1} Hz  ch={2}" -f $Info.ACodec, $Info.ASampleRate, $Info.AChannels)
    Write-Host ("  Bitrate   : ~{0} kb/s" -f $Info.BitrateKbps)
    Write-Host ""
}
function Get-UpscaleTarget {
    param(
        [Parameter(Mandatory)]$SrcInfo
    )

    # Keep it simple: user chooses a target "profile"
$choices = @(
    [pscustomobject]@{ Key="1";  Preset="av1_4k_mkv";       Name="UHD-Modern: 4K AV1 + Opus in MKV (best size)";
        W=3840; H=2160; Container="mkv"; VCodec="av1";  ACodec="opus"; PixFmt="yuv420p" },

    [pscustomobject]@{ Key="2";  Preset="av1_4k_mp4";       Name="UHD-Modern: 4K AV1 + AAC in MP4 (best size + compatibility)";
        W=3840; H=2160; Container="mp4"; VCodec="av1";  ACodec="aac";  PixFmt="yuv420p" },

    [pscustomobject]@{ Key="3";  Preset="hevc_4k_mp4";      Name="UHD-Standard: 4K HEVC CRF 22 in MP4 (smaller, good quality)";
        W=3840; H=2160; Container="mp4"; VCodec="hevc"; ACodec="aac";  PixFmt="yuv420p" },

    [pscustomobject]@{ Key="4";  Preset="h264_4k_mp4";      Name="UHD-Legacy: 4K H.264 CRF 20 in MP4 (bigger, compatible)";
        W=3840; H=2160; Container="mp4"; VCodec="h264"; ACodec="aac";  PixFmt="yuv420p" },

    [pscustomobject]@{ Key="5";  Preset="av1_1080p_mkv";    Name="HD-Modern: 1080p AV1 + Opus in MKV (best size)";
        W=1920; H=1080; Container="mkv"; VCodec="av1";  ACodec="opus"; PixFmt="yuv420p" },

    [pscustomobject]@{ Key="6";  Preset="av1_1080p_mp4";    Name="HD-Modern: 1080p AV1 + AAC in MP4 (best size + compatibility)";
        W=1920; H=1080; Container="mp4"; VCodec="av1";  ACodec="aac";  PixFmt="yuv420p" },

    [pscustomobject]@{ Key="7";  Preset="hevc_1080p_mp4";   Name="HD-Standard: 1080p H.265 CRF 20 in MP4 (smaller file)";
        W=1920; H=1080; Container="mp4"; VCodec="hevc"; ACodec="aac";  PixFmt="yuv420p" },

    [pscustomobject]@{ Key="8";  Preset="h264_1080p_mp4";   Name="HD-Legacy: 1080p H.264 CRF 18 in MP4 (high quality)";
        W=1920; H=1080; Container="mp4"; VCodec="h264"; ACodec="aac";  PixFmt="yuv420p" },

    [pscustomobject]@{ Key="9";  Preset="prores_4k_mov";    Name="Archival-UHD: 4K ProRes 422 HQ + PCM in MOV (editing)";
        W=3840; H=2160; Container="mov"; VCodec="prores"; ACodec="pcm"; PixFmt="yuv422p10le" },

    [pscustomobject]@{ Key="10"; Preset="ffv1_4k_mkv";      Name="Archival-UHD: 4K FFV1 + FLAC in MKV (huge lossless master)";
        W=3840; H=2160; Container="mkv"; VCodec="ffv1";  ACodec="flac"; PixFmt="yuv420p" },

    [pscustomobject]@{ Key="11"; Preset="prores_1080p_mov"; Name="Archival-HD: 1080p ProRes 422 HQ + PCM in MOV (editing)";
        W=1920; H=1080; Container="mov"; VCodec="prores"; ACodec="pcm"; PixFmt="yuv422p10le" },

    [pscustomobject]@{ Key="12"; Preset="ffv1_1080p_mkv";   Name="Archival-HD: 1080p FFV1 + FLAC in MKV (lossless master)";
        W=1920; H=1080; Container="mkv"; VCodec="ffv1";  ACodec="flac"; PixFmt="yuv420p" },

    [pscustomobject]@{ Key="13"; Preset="match_source";     Name="Match source (no resize), just rebuild (FFV1+FLAC MKV)";
        W=$SrcInfo.Width; H=$SrcInfo.Height; Container="mkv"; VCodec="ffv1"; ACodec="flac"; PixFmt="yuv420p" }
)



    Write-Host ""
    Write-Host "[Target Output Presets]"
    foreach ($c in $choices) {
        Write-Host ("  [{0}] {1}" -f $c.Key, $c.Name)
    }

    $sel  = Read-Host "Select 1-13"
    $pick = $choices | Where-Object { $_.Key -eq $sel } | Select-Object -First 1
    if (-not $pick) { throw "Invalid selection: $sel" }

    if ($pick.Preset -eq "match_source") {
        $preset = [pscustomobject]@{
            Name      = "match_source"
            Ext       = "mkv"
            VCodec    = "ffv1"
            VArgs     = @("-level","3","-g","1","-pix_fmt","yuv420p")
            ACodec    = "flac"
            AArgs     = @()
            Scale     = $null
            ExtraMaps = @()
        }
    } else {
        $preset = Get-OutputPresetConfig -PresetName $pick.Preset
        $preset.Scale = "$($pick.W):$($pick.H)"

        if ($preset.Ext -ne $pick.Container) {
            Write-Warning "Menu container '$($pick.Container)' != preset extension '$($preset.Ext)'. Using preset '$($preset.Ext)'."
        }
    }

    # Return BOTH, so caller can update OutputPreset string + use preset object
    return [pscustomobject]@{
        Pick   = $pick
        Preset = $preset
    }
}
function Get-PngCount($dir) {
    if (-not (Test-Path -LiteralPath $dir)) { return 0 }
    return (Get-ChildItem -LiteralPath $dir -Filter "frame_*.png" -File -ErrorAction SilentlyContinue | Measure-Object).Count
}
function Find-ExistingProjects([string]$Root = $BaseDir) {
    $projects = @()

    Get-ChildItem -Directory (Join-Path $Root "_ffv1_work_gpu*") -ErrorAction SilentlyContinue |  ForEach-Object {
        if ($_.Name -match 'gpu(\d+)') {
            $gpu = [int]$Matches[1]
            Get-ChildItem -Directory $_.FullName -ErrorAction SilentlyContinue | ForEach-Object {
                $job = $_
                $frames   = Join-Path $job.FullName "frames"
                $upscaled = Join-Path $job.FullName "upscaled"

                $fCount = Get-PngCount $frames
                $uCount = Get-PngCount $upscaled

                if ($fCount -gt 0 -and $uCount -lt $fCount) {
                    $pct = [math]::Round(($uCount / $fCount) * 100, 2)
                    $projects += [pscustomobject]@{
                        Name          = $job.Name
                        GpuIndex      = $gpu
                        Frames        = $fCount
                        Upscaled      = $uCount
                        Percent       = $pct
                        JobDir        = $job.FullName
                    }
                }
            }
        }
    }

    return $projects
}


# ---------------- DECISION: DryRun / Resume selection ----------------

$existing = Find-ExistingProjects

# Dry-run: report unfinished jobs and exit (no prompts)
if ($DryRun) {
    if (-not $existing -or $existing.Count -eq 0) {
        Write-Host "No existing unfinished projects detected."
    } else {
        foreach ($p in $existing) {
            $frames   = Join-Path $p.JobDir "frames"
            $upscaled = Join-Path $p.JobDir "upscaled"
            Show-DryRunResume $p.Name $frames $upscaled
        }
    }
    Write-Host "`n[DryRun] No actions performed."
    exit 0
}

# Helper: prompt user to pick a project index (only when needed)
function Select-ResumeProject {
    param([Parameter(Mandatory)]$Projects)

    Write-Host ""
    Write-Host "Detected unfinished projects:"
    for ($i = 0; $i -lt $Projects.Count; $i++) {
        $p = $Projects[$i]
        Write-Host ("[{0}] {1} (gpu{2}, {3}%)" -f ($i + 1), $p.Name, $p.GpuIndex, $p.Percent)
    }

    $sel = Read-Host "Select project number to resume, or press Enter to start new"
    if ([string]::IsNullOrWhiteSpace($sel)) { return $null }

    if ($sel -notmatch '^\d+$') { throw "Invalid selection: $sel" }
    $idx = [int]$sel - 1
    if ($idx -lt 0 -or $idx -ge $Projects.Count) { throw "Selection out of range: $sel" }

    return $Projects[$idx]
}

# If user asked to resume, pick a job (prompt only if needed)
$ResumeJobName = $null

$existing = Find-ExistingProjects -Root $BaseDir

if ($Resume) {
    if ($existing.Count -eq 0) {
        Write-Warning "Resume requested (-Resume) but no unfinished projects were found. Starting a new run."
    }
    elseif ($existing.Count -eq 1) {
        $p = $existing[0]
        $ResumeJobName = $p.Name
        $GpuIndex      = $p.GpuIndex
        $GpuTag        = "gpu$GpuIndex"

        Write-Host ""
        Write-Host "Resuming unfinished project automatically:"
        Write-Host "  Project : $ResumeJobName"
        Write-Host "  GPU     : $GpuTag"
        Write-Host "  Progress: $($p.Percent)%"
    }
    else {
        $p = Select-ResumeProject -Projects $existing
        if ($p) {
            $ResumeJobName = $p.Name
            $GpuIndex      = $p.GpuIndex
            $GpuTag        = "gpu$GpuIndex"
        } else {
            # user chose to start new, so turn off resume
            $Resume = $false
        }
    }
}
else {
    # Not in resume mode: if unfinished jobs exist, offer to resume (prompt once)
    if ($existing.Count -eq 1) {
        $p = $existing[0]
        Write-Host ""
        Write-Host "Detected an unfinished project:"
        Write-Host "  Project : $($p.Name)"
        Write-Host "  GPU     : gpu$($p.GpuIndex)"
        Write-Host "  Progress: $($p.Percent)%"

        $resp = Read-Host "Resume it? (Y/N)"
        if ($resp.Trim().ToUpper() -eq "Y") {
            $Resume        = $true
            $ResumeJobName = $p.Name
            $GpuIndex      = $p.GpuIndex
            $GpuTag        = "gpu$GpuIndex"
        }
    }
    elseif ($existing.Count -gt 1) {
        $p = Select-ResumeProject -Projects $existing
        if ($p) {
            $Resume        = $true
            $ResumeJobName = $p.Name
            $GpuIndex      = $p.GpuIndex
            $GpuTag        = "gpu$GpuIndex"
        }
    }
}

# If we are resuming and user did NOT explicitly set -NoDelete, keep temp dirs by default.
if ($Resume -and -not $PSBoundParameters.ContainsKey('NoDelete')) {
    $NoDelete = $true
}

# If resuming, enforce busy/lock warning on the chosen GPU (still gives user a chance to abort)
if ($Resume -and $GpuIndex -ge 0) {
    Test-AnyGpuBusyOrLocked -BusyThresholdPercent $BusyThresholdPercent
}

# ---------------- END DECISION BLOCK ----------------


function Get-VideoControllers {
    # Adapter names from Windows
    Get-CimInstance Win32_VideoController |
      Select-Object Name, AdapterCompatibility, VideoProcessor, PNPDeviceID, AdapterRAM
}

function Show-DryRunResume($jobName, $framesDir, $upscaledDir) {
    $frameCount    = Get-PngCount $framesDir
    $upscaledCount = Get-PngCount $upscaledDir

    Write-Host ""
    Write-Host "[DryRun] $jobName"

    if ($frameCount -eq 0) {
        Write-Host "  ERROR: No extracted frames"
        return
    }

    $missing = $frameCount - $upscaledCount
    $pct = [math]::Round(($upscaledCount / $frameCount) * 100, 2)

    Write-Host "  Frames extracted : $frameCount"
    Write-Host "  Frames upscaled  : $upscaledCount"
    Write-Host "  Missing frames   : $missing"
    Write-Host "  Completion       : $pct%"

    if ($upscaledCount -eq $frameCount) {
        Write-Host "  Action           : SKIP (already complete)"
    }
    elseif ($upscaledCount -lt $frameCount) {
        Write-Host "  Action           : RESUME (upscale missing frames)"
    }
    else {
        Write-Host "  Action           : ERROR (more upscaled than frames)"
    }
}


function Get-Fps($file) {
    $r = & $script:FfprobeExe -v error -select_streams v:0 -show_entries stream=r_frame_rate -of default=nw=1 "file:$file"
    $r = $r -replace "r_frame_rate=", ""
    if (-not $r) { return 30.0 }
    if ($r -match "/") {
        $p = $r.Split("/")
        if ($p.Count -eq 2 -and [double]$p[1] -ne 0) { return [double]$p[0] / [double]$p[1] }
    }
    return [double]$r
}
function Get-FrameMap($dir) {
    # Returns a HashSet-like dictionary of "frame_000001.png" -> $true
    $map = @{}
    if (-not (Test-Path $dir)) { return $map }
    Get-ChildItem -Path $dir -Filter "frame_*.png" -File -ErrorAction SilentlyContinue | ForEach-Object {
        $map[$_.Name] = $true
    }
    return $map
}

function Update-UpscaledFrames($framesDir, $upscaledDir, $todoDir, $UpscaylExe, $ModelName, $ScaleFactor, $GpuIdx, $GpuTagText, $VideoName) {
    $frameFiles = Get-ChildItem -Path $framesDir -Filter "frame_*.png" -File -ErrorAction Stop
    if (-not $frameFiles -or $frameFiles.Count -eq 0) { throw "No extracted frames found in $framesDir" }

    $upMap = Get-FrameMap $upscaledDir

    # Build todo dir
    Remove-Item -Recurse -Force $todoDir -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force $todoDir | Out-Null

    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($f in $frameFiles) {
        if (-not $upMap.ContainsKey($f.Name)) {
            $missing.Add($f.FullName)
        }
    }

    if ($missing.Count -eq 0) {
        Write-Host "[Resume] Upscaled set is complete -> no missing frames"
        return
    }

    Write-Host "[Resume] Missing upscaled frames: $($missing.Count) -> upscaling ONLY missing frames"

    # Create hardlinks for missing frames into todo folder (fast, no duplicate data)
    foreach ($src in $missing) {
        $leaf = Split-Path $src -Leaf
        $dst  = Join-Path $todoDir $leaf
        if ((Get-Item $src).PSDrive.Name -ne (Get-Item $todoDir).PSDrive.Name) {
            Copy-Item $src $dst
        } else {
            New-Item -ItemType HardLink -Path $dst -Target $src | Out-Null
        }

    }

    # Run Upscayl on todo folder; output directly into upscaled folder
    Write-Host "[Upscale Missing | $GpuTagText] $VideoName"
    $gpuArgs = @()
    if ($GpuIdx -ge 0) { $gpuArgs = @('-g', $GpuIdx) }

    & $UpscaylExe -i $todoDir -o $upscaledDir -n $ModelName -ScaleFactor $scaleForThis @gpuArgs

    if ($LASTEXITCODE -ne 0) { throw "Upscayl failed while processing missing frames for $VideoName" }
}


# ---------------- MAIN LOOP ----------------
if (-not (Test-Path -LiteralPath $script:FfmpegExe))  { throw "ffmpeg not found: $script:FfmpegExe" }
if (-not (Test-Path -LiteralPath $script:FfprobeExe)) { throw "ffprobe not found: $script:FfprobeExe" }
if (-not (Test-Path -LiteralPath $script:UpscaylExe)) { throw "upscayl not found: $script:UpscaylExe" }


New-Item -ItemType Directory -Force $WorkDir   | Out-Null
New-Item -ItemType Directory -Force $OutputDir | Out-Null

$inputFiles = Get-InputVideoFiles -InputDir $InputDir -Extensions $Extensions

# If OutputPreset not explicitly provided, show source info + ask target
$preset = $null

if (-not $PSBoundParameters.ContainsKey("OutputPreset")) {
    $first = $inputFiles | Select-Object -First 1
    if ($first) {
        $info = Get-VideoInfo -Path $first.FullName
        Show-VideoInfo -Info $info

        $choice = Get-UpscaleTarget -SrcInfo $info
        $OutputPreset = $choice.Pick.Preset   # for naming / logging
        $preset       = $choice.Preset        # the actual ffmpeg config
    }
}

if (-not $preset) {
    $preset = Get-OutputPresetConfig -PresetName $OutputPreset
}

$targetForSummary = Get-TargetWHFromPreset -Preset $preset -SrcInfo (Get-VideoInfo -Path ($inputFiles | Select-Object -First 1).FullName)

Write-Host ("Output    : {0} ({1}x{2})  v={3}  a={4}  pixfmt={5}" -f `
    $preset.Ext, `
    $targetForSummary.W, $targetForSummary.H, `
    $preset.VCodec, $preset.ACodec, `
    (Get-PixFmtFromVArgs -VArgs $preset.VArgs)
)

# ---------------- SUMMARY ----------------
Write-Host ""
Write-Host "Input Dir : $InputDir"
Write-Host "GPU       : $GpuIndex ($GpuTag)"
Write-Host "Resume    : $Resume"
Write-Host "NoDelete  : $NoDelete"
Write-Host "Work Dir  : $WorkDir"
Write-Host "Output Dir: $OutputDir"
Write-Host ""

Write-Host ("[Scan] Found {0} input file(s) in {1}" -f $inputFiles.Count, $InputDir)
if ($inputFiles.Count -eq 0) {
    Write-Warning "No input files matched extension filter. If your file is .webm, ensure it ends with .webm and is inside the input directory."
}

if ($Resume -and $ResumeJobName) {
    $inputFiles = $inputFiles | Where-Object { $_.BaseName -eq $ResumeJobName }
}

foreach ($file in $inputFiles) {

    $inFile    = $file.FullName
    $srcInfo = Get-VideoInfo -Path $inFile
    $target  = Get-TargetWHFromPreset -Preset $preset -SrcInfo $srcInfo

    # If preset is match_source, keep user Scale (this is “enhance/sharpen only” mode).
    # Otherwise auto-pick scale to avoid generating huge frames.
    $scaleForThis = $Scale
    if ($target.Mode -ne "match_source") {
        $scaleForThis = Get-UpscaylScaleForTarget `
            -SrcW $srcInfo.Width -SrcH $srcInfo.Height `
            -TargetW $target.W -TargetH $target.H `
            -MaxScale 4 -MinScale 1
    }

    if ($scaleForThis -ne $Scale) {
        Write-Host ("[Scale] Auto-adjust: user={0}x -> this file={1}x (src {2}x{3} -> target {4}x{5})" -f `
            $Scale, $scaleForThis, $srcInfo.Width, $srcInfo.Height, $target.W, $target.H)
    }

    Write-WarningIfNoResolutionGain -SrcInfo $srcInfo -TargetW $target.W -TargetH $target.H -UpscaylScale $Scale
    # If you want it interactive:
    # Write-WarningIfNoResolutionGain -SrcInfo $srcInfo -TargetW $target.W -TargetH $target.H -UpscaylScale $Scale -PromptToContinue

    $base      = $file.BaseName
    $baseSafe  = Get-SafeName $base

    $jobDir   = Join-Path $WorkDir $baseSafe
    $frames   = Join-Path $jobDir "frames"
    $upscaled = Join-Path $jobDir "upscaled"
    $todo     = Join-Path $jobDir "_todo_missing"

    $outFile = Get-OutputPathForPreset -OutputDir $OutputDir -BaseName $baseSafe -GpuTag $GpuTag -Preset $preset

    if (Test-Path $outFile) {
        if (Test-Path $outFile -and -not $Force -and (Test-OutputValid -file $outFile)) {
            Write-Host "[SKIP] Valid output exists: $($file.Name)"
            continue
        }
        if (-not $Force) {
            Write-Warning "Output exists but invalid; will rebuild: $outFile"
        } else {
            Write-Warning "Force enabled; overwriting: $outFile"
        }
    }

    # If not resuming, wipe the whole job dir
    if (-not $Resume) {
        Remove-Item -Recurse -Force $jobDir -ErrorAction SilentlyContinue
    }

    New-Item -ItemType Directory -Force $frames   | Out-Null
    New-Item -ItemType Directory -Force $upscaled | Out-Null

    if (-not (Test-Path -LiteralPath $inFile)) { throw "Input missing: $inFile" }
    Write-Host "  ffmpeg input: $inFile"


    # 1) Extract frames (skip if resuming and frames exist)
    if ($Resume -and (Get-PngCount $frames) -gt 0) {
        Write-Host "[Resume] Frames exist -> skipping extract: $($file.Name)"
    } else {
        Write-Host "`n[Extract] $($file.Name)"
        Remove-Item "$frames\*.png" -Force -ErrorAction SilentlyContinue
        & $script:FfmpegExe -y -hide_banner -loglevel warning -stats `
        -i "file:$inFile" `
        "$frames\frame_%06d.png"
        Write-Host "  frames dir: $frames"
        Get-ChildItem -LiteralPath $frames -Filter "frame_*.png" -File -ErrorAction SilentlyContinue | Select-Object -First 5 FullName

        if ($LASTEXITCODE -ne 0) { throw "Frame extraction failed: $($file.Name)" }
    }

    # 2) Upscale frames
    $frameCount    = Get-PngCount $frames
    $upscaledCount = Get-PngCount $upscaled

    if ($frameCount -le 0) { throw "No frames found after extraction for $($file.Name)" }

    if ($Resume) {
        if ($upscaledCount -eq $frameCount) {
            Write-Host "[Resume] Upscale complete ($upscaledCount/$frameCount): $($file.Name)"
        } else {
            # Smart resume: only upscale missing frames (keeps what already succeeded)
            Update-UpscaledFrames -framesDir $frames -upscaledDir $upscaled -todoDir $todo `
                -UpscaylExe $Upscayl -ModelName $Model -ScaleFactor $Scale -GpuIdx $GpuIndex `
                -GpuTagText $GpuTag -VideoName $file.Name

            # Recount after filling missing
            $upscaledCount2 = Get-PngCount $upscaled
            if ($upscaledCount2 -ne $frameCount) {
                throw "Upscaled set still incomplete after resume ($upscaledCount2/$frameCount) for $($file.Name)"
            }
        }
    } else {
        # Fresh run: clear upscaled and do full upscale
        Write-Host "[Upscale Full | $GpuTag] $($file.Name)"
        Remove-Item "$upscaled\*.png" -Force -ErrorAction SilentlyContinue
        $gpuArgs = @()
        if ($GpuIndex -ge 0) { $gpuArgs = @('-g', $GpuIndex) }  # only when user picked
        & $Upscayl -i $frames -o $upscaled -n $Model -s $scaleForThis @gpuArgs
        if ($LASTEXITCODE -ne 0) { throw "Upscayl failed: $($file.Name)" }

        $upscaledCount2 = Get-PngCount $upscaled
        if ($upscaledCount2 -ne $frameCount) {
            throw "Upscaled set incomplete after full run ($upscaledCount2/$frameCount) for $($file.Name)"
        }
    }

    # 2.5 Normalize frames ONLY if mismatch vs target
    if ($null -eq $srcInfo) {
    $srcInfo = Get-VideoInfo -Path $inFile
    }

    $target  = Get-TargetWHFromPreset -Preset $preset -SrcInfo $srcInfo

    $normalizedDir = Join-Path $jobDir "frames_normalized"
    $framePattern  = "frame_%06d.png"

    $finalFramesDir = Test-FramesMatchTarget `
        -FfmpegExe $script:FfmpegExe `
        -FfprobeExe $script:FfprobeExe `
        -UpscaledFramesDir $upscaled `
        -TargetW $target.W -TargetH $target.H `
        -NormalizedFramesDir $normalizedDir `
        -FramePattern $framePattern `
        -OutExt "png" `
        -AspectPolicy "Pad" `
        -DryRun:$DryRun

    $rebuildInputPattern = Join-Path $finalFramesDir $framePattern

    # 3) Rebuild FFV1
    $fps = Get-Fps $inFile
    Write-Host "[Rebuild | $($preset.Name) | $GpuTag] $($file.Name)  FPS=$fps"

    # Build optional scale filter (only if preset has Scale)
    # Build optional scale filter:
    # If we normalized (finalFramesDir != upscaled), the frames already match target -> no scale.
    $vfArgs = @()
    if ($preset.Scale -and ($finalFramesDir -eq $upscaled)) {
        $vfArgs = @("-vf", "scale=$($preset.Scale):flags=lanczos")
    }

    # Overwrite behavior
    $yOrN = @()
    if ($Force) { $yOrN = @("-y") } else { $yOrN = @("-n") }

    # Encode
    # optional audio map
    & $script:FfmpegExe @yOrN -v error -stats `
        -framerate $fps -i "$rebuildInputPattern" `
        -i "$inFile" `
        -map 0:v -map 1:a:0? `
        @vfArgs `
        -c:v $preset.VCodec @($preset.VArgs) `
        -c:a $preset.ACodec @($preset.AArgs) `
        "$outFile"

    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Rebuild failed; keeping temp files: $jobDir"
        throw "Rebuild failed: $($file.Name)"
    }

    # Validation: require video; audio is optional for some sources
    if (-not (Test-Path -LiteralPath $outFile)) {
        Write-Warning "Output missing; keeping temp files: $jobDir"
        throw "Output missing: $outFile"
    }

    if (-not (Test-OutputValid -file $outFile)) {
        Write-Warning "Output failed validation; keeping temp files: $jobDir"
        throw "Invalid output: $outFile"
    }


    Write-Host "[OK] Created: $(Split-Path $outFile -Leaf)"

    # 4) Cleanup (only after verified output)
    if ($NoDelete) {
        Write-Host "[NoDelete] Keeping temp files: $jobDir"
    } else {
        Remove-Item -Recurse -Force $jobDir -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "All done. FFV1 masters are in: $OutputDir"
