<#
.SYNOPSIS
    Rebuild InfoPath .xsn form templates from the folders produced by "Unpack XSN.ps1", keeping
    the cabinet structure of the original so nothing is lost.

.DESCRIPTION
    Reads the _xsn-cab.json sidecar written by Unpack XSN.ps1 and rebuilds the cabinet with the
    same file order, the same compression algorithm (MSZIP or LZX:n), the same per-file DOS
    timestamps, the same attribute flags, and the same cabinet set id / index. Files you added to
    the folder since unpacking are appended; files you deleted are reported and skipped.

    Every rebuilt .xsn is then verified before it is handed back:
      1. the new cabinet is re-parsed and its file table compared with the sidecar;
      2. the new cabinet is extracted to a temp folder and every file is SHA256-compared with the
         file it was built from - a byte-level proof that no content changed;
      3. if the folder is unedited, the new .xsn is compared with the original's SHA256, which on
         these forms normally comes out byte-for-byte identical.

    No parameters. Run it from the folder that holds your .unpacked folders (or double-click the
    .bat). Output goes to the _Packed subfolder - the original .xsn files are never touched.

.NOTES
    Target: Windows PowerShell 5.1. Requires makecab.exe and expand.exe (both built into Windows).
#>

$ErrorActionPreference = 'Stop'

# Resolve script root AFTER any binding (never in a param default), so the script works when
# launched via the .bat (powershell.exe -File ...) as well as directly.
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Path }
if (-not $scriptRoot) { $scriptRoot = (Get-Location).Path }

$SidecarName = '_xsn-cab.json'
$OutDirName  = '_Packed'
$LogPath     = Join-Path $scriptRoot '_Pack-Log.txt'

# Files that live in an unpacked folder but are not part of the form.
$NonFormFiles = @($SidecarName, '_Unpack-Log.txt', '_Pack-Log.txt', 'Thumbs.db', 'desktop.ini')

$script:Issues = @()

function Log-Issue {
    param([string]$Form, [string]$Stage, [string]$Message, [string]$Severity = 'Error')
    $script:Issues += [pscustomobject]@{ Form = $Form; Stage = $Stage; Severity = $Severity; Message = $Message }
    $color = if ($Severity -eq 'Error') { 'Red' } else { 'Yellow' }
    Write-Host ("    [{0}] {1}: {2}" -f $Severity, $Stage, $Message) -ForegroundColor $color
}

# ============================================================================================
#  Cabinet structure reader (MS-CAB) - same parser as Unpack XSN.ps1, plus the byte offset of
#  each record so the fields makecab cannot express can be patched back in afterwards.
# ============================================================================================

function Get-CabString {
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
    param([int]$DosDate, [int]$DosTime)
    $year = (($DosDate -shr 9) -band 0x7F) + 1980
    $month = ($DosDate -shr 5) -band 0x0F
    $day = $DosDate -band 0x1F
    $hour = ($DosTime -shr 11) -band 0x1F
    $min = ($DosTime -shr 5) -band 0x3F
    $sec = ($DosTime -band 0x1F) * 2
    if ($month -lt 1 -or $month -gt 12 -or $day -lt 1 -or $day -gt 31 -or $hour -gt 23 -or $min -gt 59 -or $sec -gt 59) { return $null }
    try { return (Get-Date -Year $year -Month $month -Day $day -Hour $hour -Minute $min -Second $sec -Millisecond 0) }
    catch { return $null }
}

function Get-CompressionName {
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
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 36) { throw "File is only $($bytes.Length) bytes - not a cabinet." }

    $base = -1
    $limit = [Math]::Min($bytes.Length - 36, 65536)
    for ($i = 0; $i -le $limit; $i++) {
        if ($bytes[$i] -eq 0x4D -and $bytes[$i + 1] -eq 0x53 -and $bytes[$i + 2] -eq 0x43 -and $bytes[$i + 3] -eq 0x46) { $base = $i; break }
    }
    if ($base -lt 0) { throw 'No MSCF cabinet signature found.' }

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
    }

    $off = $base + 36
    $cbCFFolder = 0
    if ($hdr.Flags -band 0x0004) {
        $cbCFHeader = [int][BitConverter]::ToUInt16($bytes, $off)
        $cbCFFolder = [int]$bytes[$off + 2]
        $off += 4 + $cbCFHeader
    }
    if ($hdr.Flags -band 0x0001) { $r = [ref]$off; [void](Get-CabString $bytes $r $false); [void](Get-CabString $bytes $r $false); $off = $r.Value }
    if ($hdr.Flags -band 0x0002) { $r = [ref]$off; [void](Get-CabString $bytes $r $false); [void](Get-CabString $bytes $r $false); $off = $r.Value }

    $folders = @()
    for ($f = 0; $f -lt $hdr.FolderCount; $f++) {
        $typeCompress = [int][BitConverter]::ToUInt16($bytes, $off + 6)
        $folders += [pscustomobject]@{
            Index = $f; DataBlocks = [int][BitConverter]::ToUInt16($bytes, $off + 4)
            TypeCompress = $typeCompress; Compression = (Get-CompressionName $typeCompress)
        }
        $off += 8 + $cbCFFolder
    }

    $files = @()
    $o = $base + [int]$hdr.CoffFiles
    for ($n = 0; $n -lt $hdr.FileCount; $n++) {
        $entry   = $o
        $cbFile  = [BitConverter]::ToUInt32($bytes, $o)
        $iFolder = [int][BitConverter]::ToUInt16($bytes, $o + 8)
        $dosDate = [int][BitConverter]::ToUInt16($bytes, $o + 10)
        $dosTime = [int][BitConverter]::ToUInt16($bytes, $o + 12)
        $attribs = [int][BitConverter]::ToUInt16($bytes, $o + 14)
        $o += 16
        $r = [ref]$o
        $name = Get-CabString -Bytes $bytes -Offset $r -Utf8 ([bool]($attribs -band 0x80))
        $o = $r.Value
        $files += [pscustomobject]@{
            Order = $n; Name = $name; Size = [long]$cbFile; FolderIndex = $iFolder
            DosDate = $dosDate; DosTime = $dosTime; Attribs = $attribs
            AttribsOffset = $entry + 14
        }
    }

    return [pscustomobject]@{
        Bytes = $bytes; Base = $base; Header = $hdr; Folders = $folders; Files = $files
        SetIdOffset = $base + 32; CabinetIndexOffset = $base + 34
    }
}

function Expand-CabTo {
    # Used only to verify what we just built. expand.exe insists on a real cabinet, so work from
    # a temp .cab copy.
    param([string]$SrcPath, [string]$DestDir, [string]$Tag, [string]$ProbeFile)
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ($Tag + '_' + ([System.IO.Path]::GetRandomFileName().Replace('.', '')) + '.cab')
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
#  Build
# ============================================================================================

function Get-MakeCabSettings {
    <#
        Translate the recorded typeCompress word into the two .Set directives makecab understands.
        makecab cannot produce Quantum (nothing has written Quantum cabinets since the 1990s and
        InfoPath never did) - fall back to LZX:21 and say so.
    #>
    param([int]$TypeCompress, [string]$FormName)
    switch ($TypeCompress -band 0x000F) {
        0 { return @{ Type = 'NONE'; Memory = 0 } }
        1 { return @{ Type = 'MSZIP'; Memory = 0 } }
        3 {
            $win = ($TypeCompress -shr 8) -band 0x1F
            if ($win -lt 15 -or $win -gt 21) { $win = 21 }
            return @{ Type = 'LZX'; Memory = $win }
        }
        default {
            Log-Issue $FormName 'build' ("Original used {0}, which makecab cannot write. Rebuilding as LZX:21 - file content is unaffected." -f (Get-CompressionName $TypeCompress)) 'Warning'
            return @{ Type = 'LZX'; Memory = 21 }
        }
    }
}

function Invoke-MakeCab {
    # Build one cabinet from an ordered list of [pscustomobject]@{ Path; Name; Folder }.
    # Returns the path of the produced cabinet, or throws.
    param([object[]]$Entries, [hashtable]$Compression, [string]$WorkDir, [string]$CabName)

    $ddfLines = @(
        '.OPTION EXPLICIT'
        ('.Set CabinetNameTemplate="{0}"' -f $CabName)
        ('.Set DiskDirectoryTemplate="{0}"' -f $WorkDir)
        '.Set Cabinet=on'
        ('.Set Compress={0}' -f $(if ($Compression.Type -eq 'NONE') { 'off' } else { 'on' }))
    )
    if ($Compression.Type -ne 'NONE') {
        $ddfLines += ('.Set CompressionType={0}' -f $Compression.Type)
        if ($Compression.Type -eq 'LZX') { $ddfLines += ('.Set CompressionMemory={0}' -f $Compression.Memory) }
    }
    $ddfLines += @(
        '.Set MaxDiskSize=0'          # one cabinet, no splitting
        '.Set FolderSizeThreshold=0'  # do not start a new cabinet folder on size
        '.Set MaxErrors=1'
        '.Set UniqueFiles=off'
        '.Set InfFileName=nul'
        '.Set RptFileName=nul'
    )

    $currentFolder = if ($Entries.Count -gt 0) { $Entries[0].Folder } else { 0 }
    foreach ($e in $Entries) {
        if ($e.Folder -ne $currentFolder) { $ddfLines += '.New Folder'; $currentFolder = $e.Folder }
        $ddfLines += ('"{0}" "{1}"' -f $e.Path, $e.Name)
    }

    $ddfPath = Join-Path $WorkDir 'build.ddf'
    # makecab reads the directive file as plain text - ASCII, no BOM.
    [System.IO.File]::WriteAllLines($ddfPath, $ddfLines, (New-Object System.Text.ASCIIEncoding))

    $stdout = Join-Path $WorkDir 'makecab.out'
    $p = Start-Process -FilePath 'makecab.exe' -ArgumentList @('/F', ('"{0}"' -f $ddfPath)) -WorkingDirectory $WorkDir `
                       -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdout -RedirectStandardError (Join-Path $WorkDir 'makecab.err')
    $cabPath = Join-Path $WorkDir $CabName
    if ($p.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $cabPath)) {
        $msg = ''
        foreach ($f in @($stdout, (Join-Path $WorkDir 'makecab.err'))) {
            if (Test-Path -LiteralPath $f) { $msg += ((Get-Content -LiteralPath $f -Raw) -replace '\s+', ' ') }
        }
        throw ("makecab failed (exit {0}). {1}" -f $p.ExitCode, $msg.Trim())
    }
    return $cabPath
}

function Repair-CabMetadata {
    <#
        makecab has no directive for the cabinet set id / index, and it derives the per-file
        attribute word from the file system, so a file that was flagged read-only, hidden or
        UTF8-named inside the original cabinet comes back plain. Both live in fixed-size fields,
        so patch them in place - no offset inside the cabinet moves.
    #>
    param([string]$CabPath, $Sidecar)

    $cab = Read-CabStructure -Path $CabPath
    $bytes = $cab.Bytes
    $patched = 0

    $setId = [int]$Sidecar.Cabinet.SetId
    $idx   = [int]$Sidecar.Cabinet.CabinetIndex
    if ($cab.Header.SetId -ne $setId) {
        [Array]::Copy([BitConverter]::GetBytes([uint16]$setId), 0, $bytes, $cab.SetIdOffset, 2); $patched++
    }
    if ($cab.Header.CabinetIndex -ne $idx) {
        [Array]::Copy([BitConverter]::GetBytes([uint16]$idx), 0, $bytes, $cab.CabinetIndexOffset, 2); $patched++
    }

    $want = @{}
    foreach ($f in $Sidecar.Files) { $want[$f.Name] = [int]$f.Attribs }
    foreach ($f in $cab.Files) {
        if (-not $want.ContainsKey($f.Name)) { continue }
        $target = $want[$f.Name]
        # _A_NAME_IS_UTF (0x80) says the name bytes are UTF-8. For a pure-ASCII name the two
        # encodings are identical so the flag is safe to restore; for anything else, leave the
        # flag makecab chose or the name would be read with the wrong encoding.
        $isAscii = ($f.Name -notmatch '[^\x00-\x7F]')
        if (-not $isAscii) { $target = ($target -band 0xFF7F) -bor ($f.Attribs -band 0x80) }
        if ($f.Attribs -ne $target) {
            [Array]::Copy([BitConverter]::GetBytes([uint16]$target), 0, $bytes, $f.AttribsOffset, 2); $patched++
        }
    }

    if ($patched -gt 0) { [System.IO.File]::WriteAllBytes($CabPath, $bytes) }
    return $patched
}

# ============================================================================================
#  Main
# ============================================================================================

Write-Host ''
Write-Host '  Pack XSN' -ForegroundColor Cyan
Write-Host ('  ' + ('-' * 60)) -ForegroundColor DarkGray
Write-Host "  Folder: $scriptRoot"
Write-Host ''

$folders = @(Get-ChildItem -LiteralPath $scriptRoot -Directory -ErrorAction SilentlyContinue |
             Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName $SidecarName) } | Sort-Object Name)

if ($folders.Count -eq 0) {
    Write-Host "  No unpacked form folders here (looking for a folder containing $SidecarName)." -ForegroundColor Yellow
    Write-Host '  Run "Unpack XSN.bat" first.' -ForegroundColor Yellow
    Write-Host ''
    exit 1
}

$outDir = Join-Path $scriptRoot $OutDirName
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$built = 0
$identical = 0
foreach ($dir in $folders) {
    $formName = $dir.Name
    Write-Host ("  {0}" -f $dir.Name) -ForegroundColor White

    # ---- sidecar ---------------------------------------------------------------------------
    $sidecar = $null
    try { $sidecar = Get-Content -LiteralPath (Join-Path $dir.FullName $SidecarName) -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { Log-Issue $formName 'sidecar' "Could not read $SidecarName : $($_.Exception.Message)" 'Error'; continue }

    $targetName = $sidecar.SourceFile
    if (-not $targetName) { $targetName = $formName -replace '\.unpacked$', '' ; $targetName += '.xsn' }

    # ---- assemble the file list, in the original cabinet order -----------------------------
    $onDisk = @{}
    foreach ($f in (Get-ChildItem -LiteralPath $dir.FullName -File -ErrorAction SilentlyContinue)) {
        if ($NonFormFiles -contains $f.Name) { continue }
        $onDisk[$f.Name] = $f
    }
    $subdirs = @(Get-ChildItem -LiteralPath $dir.FullName -Directory -ErrorAction SilentlyContinue)
    foreach ($s in $subdirs) { Log-Issue $formName 'collect' "Subfolder '$($s.Name)' ignored - an .xsn cabinet is flat." 'Warning' }

    $entries = @()
    $used = @{}
    foreach ($rec in $sidecar.Files) {
        if (-not $onDisk.ContainsKey($rec.Name)) {
            Log-Issue $formName 'collect' "'$($rec.Name)' was in the original .xsn but is not in the folder - packing without it." 'Warning'
            continue
        }
        $item = $onDisk[$rec.Name]
        # Re-stamp the file with the timestamp the cabinet recorded so makecab writes the same
        # DOS date/time back. Edited files keep the original stamp too - the .xsn is a template,
        # and InfoPath compares its manifest version, not file times.
        $stamp = ConvertFrom-DosDateTime -DosDate ([int]$rec.DosDate) -DosTime ([int]$rec.DosTime)
        if ($stamp) { Set-CabTimestamp -Item $item -WallClock $stamp }
        $entries += [pscustomobject]@{ Path = $item.FullName; Name = $rec.Name; Folder = [int]$rec.Folder }
        $used[$rec.Name] = $true
        if ($item.Length -ne [long]$rec.Size) {
            Write-Host ("    edited: {0} ({1} -> {2} bytes)" -f $rec.Name, $rec.Size, $item.Length) -ForegroundColor DarkCyan
        }
    }
    $lastFolder = if ($entries.Count -gt 0) { $entries[-1].Folder } else { 0 }
    foreach ($name in ($onDisk.Keys | Sort-Object)) {
        if ($used.ContainsKey($name)) { continue }
        Log-Issue $formName 'collect' "'$name' is new since unpacking - appending it to the cabinet." 'Warning'
        $entries += [pscustomobject]@{ Path = $onDisk[$name].FullName; Name = $name; Folder = $lastFolder }
    }
    if ($entries.Count -eq 0) { Log-Issue $formName 'collect' 'Nothing to pack.' 'Error'; continue }

    # ---- build ------------------------------------------------------------------------------
    $comp = Get-MakeCabSettings -TypeCompress ([int]$sidecar.Folders[0].TypeCompress) -FormName $formName
    $compLabel = if ($comp.Type -eq 'LZX') { "LZX:$($comp.Memory)" } else { $comp.Type }
    Write-Host ("    {0} file(s), compression {1}" -f $entries.Count, $compLabel) -ForegroundColor DarkGray

    $work = Join-Path ([System.IO.Path]::GetTempPath()) ('XsnPack_' + ([System.IO.Path]::GetRandomFileName().Replace('.', '')))
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $cabPath = $null
    try {
        $cabPath = Invoke-MakeCab -Entries $entries -Compression $comp -WorkDir $work -CabName 'built.xsn'
        [void](Repair-CabMetadata -CabPath $cabPath -Sidecar $sidecar)
    } catch {
        Log-Issue $formName 'build' $_.Exception.Message 'Error'
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
        continue
    }

    # ---- verify: structure -------------------------------------------------------------------
    $ok = $true
    try {
        $new = Read-CabStructure -Path $cabPath
        if ($new.Files.Count -ne $entries.Count) {
            Log-Issue $formName 'verify' ("Rebuilt cabinet holds {0} file(s), expected {1}." -f $new.Files.Count, $entries.Count) 'Error'; $ok = $false
        }
        for ($i = 0; $i -lt [Math]::Min($new.Files.Count, $entries.Count); $i++) {
            if ($new.Files[$i].Name -ne $entries[$i].Name) {
                Log-Issue $formName 'verify' ("Order changed at slot {0}: '{1}' instead of '{2}'." -f $i, $new.Files[$i].Name, $entries[$i].Name) 'Error'; $ok = $false
            }
        }
        foreach ($rec in $sidecar.Files) {
            $m = $new.Files | Where-Object { $_.Name -eq $rec.Name } | Select-Object -First 1
            if (-not $m) { continue }
            if ($m.DosDate -ne [int]$rec.DosDate -or $m.DosTime -ne [int]$rec.DosTime) {
                Log-Issue $formName 'verify' ("Timestamp differs for '{0}'." -f $rec.Name) 'Warning'
            }
            if ($m.Attribs -ne [int]$rec.Attribs) {
                Log-Issue $formName 'verify' ("Attribute flags differ for '{0}' (0x{1:X4} vs 0x{2:X4})." -f $rec.Name, $m.Attribs, [int]$rec.Attribs) 'Warning'
            }
        }
    } catch {
        Log-Issue $formName 'verify' "Could not re-parse the rebuilt cabinet: $($_.Exception.Message)" 'Error'; $ok = $false
    }

    # ---- verify: content, byte for byte ------------------------------------------------------
    $check = Join-Path $work 'verify'
    New-Item -ItemType Directory -Path $check -Force | Out-Null
    $expandOut = Expand-CabTo -SrcPath $cabPath -DestDir $check -Tag 'XsnVerify' -ProbeFile $entries[0].Name
    $compared = 0
    foreach ($e in $entries) {
        $rt = Join-Path $check $e.Name
        if (-not (Test-Path -LiteralPath $rt)) {
            Log-Issue $formName 'verify' ("'{0}' did not come back out of the rebuilt cabinet. {1}" -f $e.Name, ($expandOut -replace '\s+', ' ').Trim()) 'Error'
            $ok = $false; continue
        }
        if ((Get-Sha256 $rt) -ne (Get-Sha256 $e.Path)) {
            Log-Issue $formName 'verify' ("Content changed for '{0}'." -f $e.Name) 'Error'; $ok = $false; continue
        }
        $compared++
    }

    # ---- deliver ------------------------------------------------------------------------------
    if ($ok) {
        $dest = Join-Path $outDir $targetName
        Copy-Item -LiteralPath $cabPath -Destination $dest -Force
        $newHash = Get-Sha256 $dest
        $same = ($sidecar.SourceSha256 -and $newHash -eq $sidecar.SourceSha256)
        if ($same) { $identical++ }
        $note = if ($same) { 'byte-identical to the original' } else { ("{0} file(s) verified byte-for-byte" -f $compared) }
        Write-Host ("    -> {0}\{1}  [{2}]" -f $OutDirName, $targetName, $note) -ForegroundColor Green
        $built++
    } else {
        Write-Host '    -> not written (verification failed, see log)' -ForegroundColor Red
    }

    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

# ============================================================================================
#  Log / summary
# ============================================================================================

$errors = @($script:Issues | Where-Object { $_.Severity -eq 'Error' })
$warns  = @($script:Issues | Where-Object { $_.Severity -eq 'Warning' })

$lines = @("Pack XSN log - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", "Folder: $scriptRoot", '')
if ($script:Issues.Count -eq 0) {
    $lines += 'No issues.'
} else {
    foreach ($i in $script:Issues) { $lines += ("[{0}] {1} | {2} | {3}" -f $i.Severity, $i.Form, $i.Stage, $i.Message) }
}
$lines | Set-Content -LiteralPath $LogPath -Encoding UTF8

Write-Host ''
Write-Host ('  ' + ('-' * 60)) -ForegroundColor DarkGray
Write-Host ("  Packed {0} of {1} form(s), {2} byte-identical to the original. Errors: {3}  Warnings: {4}" -f $built, $folders.Count, $identical, $errors.Count, $warns.Count)
Write-Host ("  Output: {0}" -f $outDir)
Write-Host ("  Log:    {0}" -f $LogPath) -ForegroundColor DarkGray
Write-Host ''

if ($errors.Count -gt 0) { exit 1 }
exit 0
