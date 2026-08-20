<#
.SYNOPSIS
    Unpack InfoPath .xsn form templates into editable folders, recording everything needed to
    rebuild the .xsn later with no data loss.

.DESCRIPTION
    An .xsn is a Microsoft cabinet (CAB). Simply renaming it to .cab and extracting loses the
    parts of the cabinet that are NOT file content: the file order, the per-file timestamps and
    attribute flags, the compression algorithm, and the cabinet set id. Repacking without those
    produces a technically valid but structurally different .xsn.

    This script parses the cabinet structure itself, extracts every file, and writes a sidecar
    manifest (_xsn-cab.json) capturing:
      * cabinet header  - version, flags, set id, cabinet index, original SHA256
      * folder table    - compression type (NONE / MSZIP / LZX:n) per cabinet folder
      * file table      - exact order, size, DOS date/time, attribute flags, SHA256

    "Pack XSN.ps1" reads that sidecar and rebuilds a matching .xsn. On unedited forms the result
    is normally byte-for-byte identical to the original.

    No parameters. Run it from the folder that holds your .xsn files (or double-click the .bat).
    Every .xsn in that folder is unpacked into a sibling "<name>.unpacked" folder.

    The .xsn files are never modified.

.NOTES
    Target: Windows PowerShell 5.1. Requires expand.exe (built into Windows).
#>

$ErrorActionPreference = 'Stop'

# Resolve script root AFTER any binding (never in a param default), so the script works when
# launched via the .bat (powershell.exe -File ...) as well as directly.
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Path }
if (-not $scriptRoot) { $scriptRoot = (Get-Location).Path }

$SidecarName = '_xsn-cab.json'
$LogPath     = Join-Path $scriptRoot '_Unpack-Log.txt'

$script:Issues = @()

function Log-Issue {
    param([string]$Form, [string]$Stage, [string]$Message, [string]$Severity = 'Error')
    $script:Issues += [pscustomobject]@{ Form = $Form; Stage = $Stage; Severity = $Severity; Message = $Message }
    $color = if ($Severity -eq 'Error') { 'Red' } else { 'Yellow' }
    Write-Host ("    [{0}] {1}: {2}" -f $Severity, $Stage, $Message) -ForegroundColor $color
}

# ============================================================================================
#  Cabinet structure reader (MS-CAB). Pure PowerShell - no external tool, nothing is modified.
# ============================================================================================

function Get-CabString {
    # Read a NUL-terminated string starting at [ref]$Offset, advancing it past the terminator.
    param([byte[]]$Bytes, [ref]$Offset, [bool]$Utf8)
    $start = $Offset.Value
    $i = $start
    while ($i -lt $Bytes.Length -and $Bytes[$i] -ne 0) { $i++ }
    $len = $i - $start
    $text = ''
    if ($len -gt 0) {
        $enc = if ($Utf8) { [System.Text.Encoding]::UTF8 } else { [System.Text.Encoding]::GetEncoding(1252) }
        $text = $enc.GetString($Bytes, $start, $len)
    }
    $Offset.Value = $i + 1
    return $text
}

function ConvertFrom-DosDateTime {
    # CAB stores local wall-clock time as packed DOS date/time (2-second resolution).
    param([int]$DosDate, [int]$DosTime)
    $year = (($DosDate -shr 9) -band 0x7F) + 1980
    $month = ($DosDate -shr 5) -band 0x0F
    $day = $DosDate -band 0x1F
    $hour = ($DosTime -shr 11) -band 0x1F
    $min = ($DosTime -shr 5) -band 0x3F
    $sec = ($DosTime -band 0x1F) * 2
    if ($month -lt 1 -or $month -gt 12 -or $day -lt 1 -or $day -gt 31 -or $hour -gt 23 -or $min -gt 59 -or $sec -gt 59) {
        return $null   # malformed stamp - keep the raw words, skip the friendly form
    }
    try { return (Get-Date -Year $year -Month $month -Day $day -Hour $hour -Minute $min -Second $sec -Millisecond 0) }
    catch { return $null }
}

function Get-CompressionName {
    # Low nibble is the algorithm; for LZX bits 8..12 carry the window size.
    param([int]$TypeCompress)
    switch ($TypeCompress -band 0x000F) {
        0 { return 'NONE' }
        1 { return 'MSZIP' }
        2 { return ('QUANTUM:{0}' -f (($TypeCompress -shr 8) -band 0x1F)) }
        3 { return ('LZX:{0}' -f (($TypeCompress -shr 8) -band 0x1F)) }
        default { return ('UNKNOWN:0x{0:X4}' -f $TypeCompress) }
    }
}

function Read-CabStructure {
    <#
        Parse a cabinet into { Base, Header, Folders, Files, Warnings }.
        Base is the offset of the MSCF signature - a few .xsn files re-saved by InfoPath carry
        leading bytes before the cabinet, and every offset inside the cabinet is relative to it.
    #>
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 36) { throw "File is only $($bytes.Length) bytes - not a cabinet." }

    # Locate MSCF. Normally at 0; scan a little way in for the re-saved variant.
    $base = -1
    $limit = [Math]::Min($bytes.Length - 36, 65536)
    for ($i = 0; $i -le $limit; $i++) {
        if ($bytes[$i] -eq 0x4D -and $bytes[$i + 1] -eq 0x53 -and $bytes[$i + 2] -eq 0x43 -and $bytes[$i + 3] -eq 0x46) { $base = $i; break }
    }
    if ($base -lt 0) { throw 'No MSCF cabinet signature found (not a CAB - possibly a ZIP or a OneDrive cloud-only stub).' }

    $warn = @()
    $hdr = [ordered]@{
        CabStartOffset = $base
        CbCabinet      = [BitConverter]::ToUInt32($bytes, $base + 8)
        CoffFiles      = [BitConverter]::ToUInt32($bytes, $base + 16)
        VersionMajor   = [int]$bytes[$base + 25]
        VersionMinor   = [int]$bytes[$base + 24]
        FolderCount    = [int][BitConverter]::ToUInt16($bytes, $base + 26)
        FileCount      = [int][BitConverter]::ToUInt16($bytes, $base + 28)
        Flags          = [int][BitConverter]::ToUInt16($bytes, $base + 30)
        SetId          = [int][BitConverter]::ToUInt16($bytes, $base + 32)
        CabinetIndex   = [int][BitConverter]::ToUInt16($bytes, $base + 34)
        ReserveHeader  = 0
        ReserveFolder  = 0
        ReserveData    = 0
        ReserveBytesHex = ''
        PrevCabinet    = ''
        NextCabinet    = ''
    }

    $off = $base + 36
    $cbCFFolder = 0
    if ($hdr.Flags -band 0x0004) {           # cfhdrRESERVE_PRESENT
        $cbCFHeader = [int][BitConverter]::ToUInt16($bytes, $off)
        $cbCFFolder = [int]$bytes[$off + 2]
        $hdr.ReserveHeader = $cbCFHeader
        $hdr.ReserveFolder = $cbCFFolder
        $hdr.ReserveData = [int]$bytes[$off + 3]
        $off += 4
        if ($cbCFHeader -gt 0) {
            $hdr.ReserveBytesHex = (($bytes[$off..($off + $cbCFHeader - 1)]) | ForEach-Object { $_.ToString('X2') }) -join ''
            $off += $cbCFHeader
        }
        $warn += 'Cabinet carries per-cabinet reserved data; repacking cannot reproduce it.'
    }
    if ($hdr.Flags -band 0x0001) {           # cfhdrPREV_CABINET
        $r = [ref]$off
        $hdr.PrevCabinet = Get-CabString -Bytes $bytes -Offset $r -Utf8 $false
        [void](Get-CabString -Bytes $bytes -Offset $r -Utf8 $false)
        $off = $r.Value
        $warn += 'Cabinet is part of a multi-cabinet set (previous cabinet declared).'
    }
    if ($hdr.Flags -band 0x0002) {           # cfhdrNEXT_CABINET
        $r = [ref]$off
        $hdr.NextCabinet = Get-CabString -Bytes $bytes -Offset $r -Utf8 $false
        [void](Get-CabString -Bytes $bytes -Offset $r -Utf8 $false)
        $off = $r.Value
        $warn += 'Cabinet is part of a multi-cabinet set (next cabinet declared).'
    }

    $folders = @()
    for ($f = 0; $f -lt $hdr.FolderCount; $f++) {
        $typeCompress = [int][BitConverter]::ToUInt16($bytes, $off + 6)
        $folders += [pscustomobject]@{
            Index        = $f
            DataOffset   = [BitConverter]::ToUInt32($bytes, $off)
            DataBlocks   = [int][BitConverter]::ToUInt16($bytes, $off + 4)
            TypeCompress = $typeCompress
            Compression  = (Get-CompressionName $typeCompress)
        }
        $off += 8 + $cbCFFolder
    }

    $files = @()
    $o = $base + [int]$hdr.CoffFiles
    for ($n = 0; $n -lt $hdr.FileCount; $n++) {
        $cbFile  = [BitConverter]::ToUInt32($bytes, $o)
        $uoff    = [BitConverter]::ToUInt32($bytes, $o + 4)
        $iFolder = [int][BitConverter]::ToUInt16($bytes, $o + 8)
        $dosDate = [int][BitConverter]::ToUInt16($bytes, $o + 10)
        $dosTime = [int][BitConverter]::ToUInt16($bytes, $o + 12)
        $attribs = [int][BitConverter]::ToUInt16($bytes, $o + 14)
        $o += 16
        $r = [ref]$o
        $name = Get-CabString -Bytes $bytes -Offset $r -Utf8 ([bool]($attribs -band 0x80))
        $o = $r.Value
        $stamp = ConvertFrom-DosDateTime -DosDate $dosDate -DosTime $dosTime
        $files += [pscustomobject]@{
            Order       = $n
            Name        = $name
            Size        = [long]$cbFile
            FolderIndex = $iFolder
            DosDate     = $dosDate
            DosTime     = $dosTime
            Modified    = if ($stamp) { $stamp.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            Attribs     = $attribs
            UncompressedOffset = [long]$uoff
        }
    }

    return [pscustomobject]@{ Base = $base; Header = $hdr; Folders = $folders; Files = $files; Warnings = $warn }
}

# ============================================================================================
#  Extraction
# ============================================================================================

function Expand-CabTo {
    # expand.exe handles every compression the cabinet format allows (MSZIP, LZX, Quantum) and is
    # built into Windows. It insists on a real cabinet, so work from a temp .cab copy - that is
    # exactly what renaming the .xsn to .cab by hand does. Shell handler is the fallback.
    param([string]$SrcPath, [string]$DestDir, [string]$FormName, [string]$ProbeFile)

    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ($FormName + '_' + ([System.IO.Path]::GetRandomFileName().Replace('.', '')) + '.cab')
    $out = ''
    $probe = if ($ProbeFile) { Join-Path $DestDir $ProbeFile } else { $null }
    try { Copy-Item -LiteralPath $SrcPath -Destination $tmp -Force } catch {}
    if (Test-Path -LiteralPath $tmp) {
        try { $out = & expand.exe $tmp '-F:*' $DestDir 2>&1 | Out-String } catch { $out = $_.Exception.Message }
    }
    if ($probe -and -not (Test-Path -LiteralPath $probe) -and (Test-Path -LiteralPath $tmp)) {
        try {
            $shell = New-Object -ComObject Shell.Application
            $s = $shell.NameSpace([string]$tmp); $d = $shell.NameSpace([string]$DestDir)
            if ($s -and $d) {
                $d.CopyHere($s.Items(), 0x14)
                for ($i = 0; $i -lt 75; $i++) { if (Test-Path -LiteralPath $probe) { break }; Start-Sleep -Milliseconds 200 }
            }
        } catch {}
    }
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath (Join-Path $DestDir ([System.IO.Path]::GetFileName($tmp))) -Force -ErrorAction SilentlyContinue
    return $out
}

function Get-Sha256 {
    param([string]$Path)
    try { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash } catch { return '' }
}

function Get-FolderState {
    # Decide whether an existing "<name>.unpacked" folder may be thrown away and re-extracted.
    #   Foreign  - it exists but this script did not make it; never touch it.
    #   Modified - it holds edits, additions or deletions; re-extracting would destroy that work.
    #   Clean    - every file still matches the hash recorded at unpack time; safe to replace.
    param([string]$Dir)
    $sidecarPath = Join-Path $Dir $SidecarName
    if (-not (Test-Path -LiteralPath $sidecarPath)) { return 'Foreign' }
    $prev = $null
    try { $prev = Get-Content -LiteralPath $sidecarPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { return 'Foreign' }
    if (-not $prev.Files) { return 'Foreign' }

    $seen = @{}
    foreach ($rec in $prev.Files) {
        $p = Join-Path $Dir $rec.Name
        if (-not (Test-Path -LiteralPath $p)) { return 'Modified' }
        if ((Get-Item -LiteralPath $p).Length -ne [long]$rec.Size) { return 'Modified' }
        if ($rec.Sha256 -and (Get-Sha256 $p) -ne $rec.Sha256) { return 'Modified' }
        $seen[$rec.Name] = $true
    }
    foreach ($f in (Get-ChildItem -LiteralPath $Dir -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -eq $SidecarName) { continue }
        if (-not $seen.ContainsKey($f.Name)) { return 'Modified' }
    }
    return 'Clean'
}

function Set-CabTimestamp {
    # Windows' cabinet tools convert file times with FileTimeToLocalFileTime, which applies
    # TODAY's daylight-saving bias rather than the bias that was in force on the file's own date.
    # .NET applies the historically correct one, so a file stamped outside the current DST season
    # would go into the cabinet an hour off. Write the UTC time that makes the Win32 conversion
    # produce exactly the wall clock the cabinet recorded.
    param([System.IO.FileInfo]$Item, [datetime]$WallClock)
    $bias = [System.TimeZoneInfo]::Local.GetUtcOffset([datetime]::Now)
    try { $Item.LastWriteTimeUtc = ([datetime]::SpecifyKind($WallClock, 'Unspecified') - $bias) } catch {}
}

# ============================================================================================
#  Main
# ============================================================================================

Write-Host ''
Write-Host '  Unpack XSN' -ForegroundColor Cyan
Write-Host ('  ' + ('-' * 60)) -ForegroundColor DarkGray
Write-Host "  Folder: $scriptRoot"
Write-Host ''

$xsnFiles = @(Get-ChildItem -LiteralPath $scriptRoot -Filter '*.xsn' -File -ErrorAction SilentlyContinue | Sort-Object Name)
if ($xsnFiles.Count -eq 0) {
    Write-Host '  No .xsn files found in this folder. Copy the script next to your forms and re-run.' -ForegroundColor Yellow
    Write-Host ''
    exit 1
}

$done = 0
foreach ($xsn in $xsnFiles) {
    $formName = [System.IO.Path]::GetFileNameWithoutExtension($xsn.Name)
    $destDir  = Join-Path $scriptRoot ($formName + '.unpacked')
    Write-Host ("  {0}" -f $xsn.Name) -ForegroundColor White

    # ---- read the cabinet structure -------------------------------------------------------
    $cab = $null
    try { $cab = Read-CabStructure -Path $xsn.FullName }
    catch {
        Log-Issue $formName 'read' $_.Exception.Message 'Error'
        continue
    }
    foreach ($w in $cab.Warnings) { Log-Issue $formName 'read' $w 'Warning' }

    $compression = ($cab.Folders | ForEach-Object { $_.Compression }) -join ', '
    Write-Host ("    {0} file(s), {1} folder(s), compression {2}" -f $cab.Files.Count, $cab.Folders.Count, $compression) -ForegroundColor DarkGray

    $dupes = @($cab.Files | Group-Object Name | Where-Object { $_.Count -gt 1 })
    foreach ($d in $dupes) { Log-Issue $formName 'read' "Cabinet holds $($d.Count) entries named '$($d.Name)'; only the last one survives on disk." 'Warning' }

    # ---- extract ---------------------------------------------------------------------------
    # Start from a clean folder so a stale file from an earlier run can never be packed back in -
    # but only ever throw away a folder that is provably untouched, so a second run of this script
    # can never destroy edits you have already made.
    if (Test-Path -LiteralPath $destDir) {
        $state = Get-FolderState -Dir $destDir
        if ($state -eq 'Foreign') {
            Log-Issue $formName 'extract' ("'{0}' already exists and was not created by this script. Move it aside and re-run." -f (Split-Path -Leaf $destDir)) 'Error'
            continue
        }
        if ($state -eq 'Modified') {
            Log-Issue $formName 'extract' ("'{0}' contains edits - left as it is so your work is not lost. Delete the folder if you want a fresh copy." -f (Split-Path -Leaf $destDir)) 'Warning'
            Write-Host '    -> skipped (already unpacked, and edited)' -ForegroundColor Yellow
            continue
        }
        try { Remove-Item -LiteralPath $destDir -Recurse -Force } catch { Log-Issue $formName 'extract' "Could not clear '$destDir': $($_.Exception.Message)" 'Error'; continue }
    }
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null

    $probe = if ($cab.Files.Count -gt 0) { $cab.Files[0].Name } else { 'manifest.xsf' }
    $expandOut = Expand-CabTo -SrcPath $xsn.FullName -DestDir $destDir -FormName $formName -ProbeFile $probe

    # ---- verify every entry landed, at the right size, and stamp the original times ---------
    $records = @()
    $missing = 0
    $badSize = 0
    foreach ($f in $cab.Files) {
        $path = Join-Path $destDir $f.Name
        $hash = ''
        if (-not (Test-Path -LiteralPath $path)) {
            $missing++
            Log-Issue $formName 'extract' "Missing after extraction: '$($f.Name)'." 'Error'
        } else {
            $item = Get-Item -LiteralPath $path
            if ($item.Length -ne $f.Size) {
                $badSize++
                Log-Issue $formName 'extract' ("Size mismatch for '{0}': cabinet says {1} bytes, extracted {2}." -f $f.Name, $f.Size, $item.Length) 'Error'
            }
            $hash = Get-Sha256 $path
            $stamp = ConvertFrom-DosDateTime -DosDate $f.DosDate -DosTime $f.DosTime
            if ($stamp) { Set-CabTimestamp -Item $item -WallClock $stamp }
        }
        $records += [ordered]@{
            Order    = $f.Order
            Name     = $f.Name
            Size     = $f.Size
            Folder   = $f.FolderIndex
            DosDate  = $f.DosDate
            DosTime  = $f.DosTime
            Modified = $f.Modified
            Attribs  = $f.Attribs
            Sha256   = $hash
        }
    }
    if ($missing -eq $cab.Files.Count -and $cab.Files.Count -gt 0) {
        Log-Issue $formName 'extract' ("Extraction produced nothing. expand.exe said: {0}" -f ($expandOut -replace '\s+', ' ').Trim()) 'Error'
        continue
    }

    # ---- sidecar ---------------------------------------------------------------------------
    $sidecar = [ordered]@{
        Tool            = 'Unpack XSN.ps1'
        FormatVersion   = 1
        UnpackedUtc     = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        SourceFile      = $xsn.Name
        SourceSize      = $xsn.Length
        SourceSha256    = (Get-Sha256 $xsn.FullName)
        Cabinet         = $cab.Header
        Folders         = @($cab.Folders | ForEach-Object { [ordered]@{ Index = $_.Index; TypeCompress = $_.TypeCompress; Compression = $_.Compression; DataBlocks = $_.DataBlocks } })
        Files           = $records
    }
    $sidecarPath = Join-Path $destDir $SidecarName
    ($sidecar | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $sidecarPath -Encoding UTF8

    $flag = if ($missing -or $badSize) { 'with errors (see log)' } else { 'OK' }
    Write-Host ("    -> {0}  [{1}]" -f (Split-Path -Leaf $destDir), $flag) -ForegroundColor Green
    $done++
}

# ============================================================================================
#  Log / summary
# ============================================================================================

$errors = @($script:Issues | Where-Object { $_.Severity -eq 'Error' })
$warns  = @($script:Issues | Where-Object { $_.Severity -eq 'Warning' })

$lines = @("Unpack XSN log - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", "Folder: $scriptRoot", '')
if ($script:Issues.Count -eq 0) {
    $lines += 'No issues.'
} else {
    foreach ($i in $script:Issues) { $lines += ("[{0}] {1} | {2} | {3}" -f $i.Severity, $i.Form, $i.Stage, $i.Message) }
}
$lines | Set-Content -LiteralPath $LogPath -Encoding UTF8

Write-Host ''
Write-Host ('  ' + ('-' * 60)) -ForegroundColor DarkGray
Write-Host ("  Unpacked {0} of {1} form(s). Errors: {2}  Warnings: {3}" -f $done, $xsnFiles.Count, $errors.Count, $warns.Count)
Write-Host ("  Log: {0}" -f $LogPath) -ForegroundColor DarkGray
Write-Host '  Edit the files in the .unpacked folders, then run "Pack XSN.bat" to rebuild the .xsn.'
Write-Host ''

if ($errors.Count -gt 0) { exit 1 }
exit 0
