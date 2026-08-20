# ================= CONFIG =================
$sourceFolder = "C:\Temp\XmlForms"        # folder containing the XML files
$outputFolder = "C:\Temp\ExtractedFiles"  # destination root folder
# ==========================================

# Ensure output root exists
if (-not (Test-Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory | Out-Null
}

function Get-InfoPathAttachment {
    param([string]$Base64String)

    try {
        $bytes = [Convert]::FromBase64String($Base64String.Trim())
    }
    catch {
        return $null  # not valid base64
    }

    # Must be long enough for the 24-byte header and match the signature
    if ($bytes.Length -lt 28) { return $null }
    if ($bytes[0] -ne 0xC7 -or $bytes[1] -ne 0x49 -or
        $bytes[2] -ne 0x46 -or $bytes[3] -ne 0x41) {
        return $null  # not an InfoPath attachment
    }

    $fileNameLenChars = [BitConverter]::ToUInt32($bytes, 20)
    $fileNameBytes    = $fileNameLenChars * 2

    $fileName = [Text.Encoding]::Unicode.GetString($bytes, 24, $fileNameBytes).TrimEnd([char]0)

    $dataStart = 24 + $fileNameBytes
    if ($dataStart -ge $bytes.Length) { return $null }

    $fileData = New-Object byte[] ($bytes.Length - $dataStart)
    [Array]::Copy($bytes, $dataStart, $fileData, 0, $fileData.Length)

    return [PSCustomObject]@{
        FileName = $fileName
        Data     = $fileData
    }
}

$xmlFiles = Get-ChildItem -Path $sourceFolder -Filter *.xml -File

foreach ($xmlFile in $xmlFiles) {

    Write-Host "Processing: $($xmlFile.Name)" -ForegroundColor Cyan

    try {
        [xml]$xml = Get-Content -Path $xmlFile.FullName -Raw
    }
    catch {
        Write-Warning "  Could not parse XML: $($xmlFile.Name). Skipping."
        continue
    }

    # Subfolder path is prepared, but NOT created yet
    $subFolder = Join-Path $outputFolder $xmlFile.BaseName
    $count = 0

    # Walk every element that has a text value, test if it's an InfoPath attachment
    $allNodes = $xml.SelectNodes("//*")
    foreach ($node in $allNodes) {

        # Skip elements that have child elements (we only want leaf text values)
        $hasChildElements = $node.ChildNodes | Where-Object { $_.NodeType -eq "Element" }
        if ($hasChildElements) { continue }

        $value = $node.InnerText
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if ($value.Length -lt 40) { continue }  # too short to be an attachment

        $attachment = Get-InfoPathAttachment -Base64String $value
        if ($null -eq $attachment) { continue }

        # Create the subfolder only when the first attachment is found
        if (-not (Test-Path $subFolder)) {
            New-Item -Path $subFolder -ItemType Directory | Out-Null
        }

        # Handle duplicate filenames inside the same XML
        $targetPath = Join-Path $subFolder $attachment.FileName
        $i = 1
        while (Test-Path $targetPath) {
            $nameNoExt  = [IO.Path]::GetFileNameWithoutExtension($attachment.FileName)
            $ext        = [IO.Path]::GetExtension($attachment.FileName)
            $targetPath = Join-Path $subFolder ("{0} ({1}){2}" -f $nameNoExt, $i, $ext)
            $i++
        }

        [IO.File]::WriteAllBytes($targetPath, $attachment.Data)
        Write-Host "  Extracted: $([IO.Path]::GetFileName($targetPath))" -ForegroundColor Green
        $count++
    }

    if ($count -eq 0) {
        Write-Host "  No attachments found, no folder created." -ForegroundColor Yellow
    }
}

Write-Host "`nDone." -ForegroundColor Cyan