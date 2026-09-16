# Excel COM finalize: recalculate formula cache, save, export PDF.
# openpyxl leaves no cached formula values, so every generated PO must pass through here.
# ASCII only on purpose: a .ps1 containing Korean must be saved UTF-8 with BOM or PS 5.1
# fails to parse it. Sheet is located by its A3 title, so no Korean argument is needed.
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Pdf = ""
)

$ErrorActionPreference = "Stop"
$xlsx = (Resolve-Path -LiteralPath $Path).Path
$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $wb = $excel.Workbooks.Open($xlsx)

    $ws = $null
    foreach ($s in $wb.Worksheets) {
        $a3 = [string]$s.Range("A3").Value2
        if ($a3 -and $a3.ToUpper().Contains("PURCHASE")) { $ws = $s; break }
    }
    if ($null -eq $ws) { throw "PURCHASE ORDER sheet not found (A3)" }
    $ws.Activate()

    $excel.CalculateFullRebuild()

    # Column-width overflow check: Excel renders '###' when a number does not fit.
    # Automated content checks cannot catch this - the PDF just shows hashes.
    $hash = @()
    for ($r = 9; $r -le 45; $r++) {
        for ($c = 1; $c -le 9; $c++) {
            $t = [string]$ws.Cells.Item($r, $c).Text
            if ($t -and $t.Contains("###")) { $hash += $ws.Cells.Item($r, $c).Address(0, 0) }
        }
    }
    if ($hash.Count -gt 0) { Write-Output ("HASH_CELLS=" + ($hash -join ",")) }
    else { Write-Output "HASH_CELLS=" }

    # Page layout probe: a vertical break means the columns no longer fit the page width
    # (the form would be sliced in half), which no print scale may trade away.
    Write-Output ("HPAGEBREAKS=" + $ws.HPageBreaks.Count)
    Write-Output ("VPAGEBREAKS=" + $ws.VPageBreaks.Count)
    $total = [double]$ws.Range("C12").Value2
    Write-Output ("AMOUNT=" + $total)

    $wb.Save()

    if ($Pdf -ne "") {
        $dir = Split-Path -Parent $Pdf
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $ws.ExportAsFixedFormat(0, $Pdf)
        Write-Output ("PDF=" + $Pdf)
    }

    $wb.Close($true)
    $wb = $null
    $excel.Quit()
    Write-Output "OK"
}
catch {
    Write-Output ("ERROR=" + $_.Exception.Message)
    try { if ($wb) { $wb.Close($false) } } catch {}
    try { if ($excel) { $excel.Quit() } } catch {}
    exit 1
}
finally {
    # Release COM handles; a leaked hidden instance keeps the file locked.
    try { if ($wb) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
    try { if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
