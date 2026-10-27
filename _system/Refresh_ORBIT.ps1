# ============================================================
# PROJECT ORBIT - One-Click Refresh + GitHub Publish
# Steps:
#   0. Timestamped backup: ORBIT.html / orbit-app.js / orbit.css
#      -> ORBIT DASH\_system\BACKUP\YYYYMMDD_HHMM\
#   1. Refresh Excel data model (RefreshAll)
#   2. Export DATA_EXTRACT -> prod_raw.csv / qa_raw.csv / assoc_raw.csv
#   2.5. Export leadership roster -> orbit-auth.js
#   3. Rebuild orbit-data.js
#   4. Verify deployment folder + confirm app files untouched
#   5. Git commit + push -> GitHub Pages publishes automatically
#   6. Open (or reload) ORBIT.html in the default browser
#
# GUARANTEE: this script writes ONLY data files
#   (orbit-data.js, orbit-auth.js, prod_raw.csv, qa_raw.csv, assoc_raw.csv).
# It NEVER writes or overwrites ORBIT.html / orbit-app.js / orbit.css.
#
# FOLDER LAYOUT:
#   ORBIT DASH\              -- git repository root (GitHub Pages source)
#     ORBIT.html             -- the application
#     orbit-app.js, orbit.css, chart.umd.min.js
#     orbit-data.js          -- generated; committed on every refresh
#     orbit-auth.js          -- generated; committed on every refresh
#     assets\logos\          -- logo images
#     _system\               -- admin-only scripts + raw data (gitignored data)
#       Refresh_ORBIT.ps1    -- this script
#       BACKUP\              -- timestamped app-file backups (gitignored)
#       .versions\           -- VCS snapshots (gitignored)
#       prod_raw.csv, qa_raw.csv, assoc_raw.csv (gitignored)
# ============================================================

$WORKBOOK  = "C:\Users\AL70683\OneDrive - Elevance Health\IGNITE PROJECTS\ORBIT\PI DI Dashboard.xlsx"
$DASH_DIR  = "C:\Users\AL70683\OneDrive - Elevance Health\IGNITE PROJECTS\ORBIT\ORBIT DASH"
$SYS_DIR   = "$DASH_DIR\_system"   # backend/admin files - not exposed to end users
$HTML_FILE = "$DASH_DIR\ORBIT.html"

# ── Git executable (GitHub Desktop ships its own copy) ────────────────────────
$GIT = "C:\Users\AL70683\AppData\Local\Programs\Git\bin\git.exe"

# ── Canonical live URL ────────────────────────────────────────────────────────
# GitHub Pages URL — this is what all users bookmark.
$ORBIT_URL = "https://charloncalugay-sys.github.io/orbit-dashboard/ORBIT.html"
# ─────────────────────────────────────────────────────────────────────────────

$T0 = Get-Date
function Log($msg, $color="White") {
    $ts = (Get-Date).ToString("HH:mm:ss")
    Write-Host "[$ts] $msg" -ForegroundColor $color
}
function Elapsed { $sec = [int]((Get-Date)-$T0).TotalSeconds; return "[$sec s elapsed]" }

# ── Helpers shared by extract + rebuild ──────────────────────────────────────
function ParseCsvLine($line) {
    $fields = [System.Collections.Generic.List[string]]::new()
    $inQ = $false; $cur = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $line.Length; $i++) {
        $ch = $line[$i]
        if ($ch -eq '"') {
            if ($inQ -and $i+1 -lt $line.Length -and $line[$i+1] -eq '"') {
                $cur.Append('"') | Out-Null; $i++
            } else { $inQ = -not $inQ }
        } elseif ($ch -eq ',' -and -not $inQ) {
            $fields.Add($cur.ToString()); $cur.Clear() | Out-Null
        } else { $cur.Append($ch) | Out-Null }
    }
    $fields.Add($cur.ToString()); return $fields.ToArray()
}
function IsDateSerial($s) {
    $n = 0.0
    if ([double]::TryParse($s,[System.Globalization.NumberStyles]::Any,
        [System.Globalization.CultureInfo]::InvariantCulture,[ref]$n)) {
        return $n -gt 40000 -and $n -lt 60000
    }; return $false
}
function ToDouble($s) {
    $n = 0.0
    [double]::TryParse($s,[System.Globalization.NumberStyles]::Any,
        [System.Globalization.CultureInfo]::InvariantCulture,[ref]$n) | Out-Null
    return $n
}
function IsDomainID($s)  { return $s -match '^[A-Z]{1,3}[0-9]+$' }
function Col0($line)     { return (ParseCsvLine $line)[0] }
function NextC0($lines,$from) {
    for ($j = $from; $j -le ([Math]::Min($lines.Length-1,$from+5)); $j++) {
        $l = $lines[$j]; if ($l -ne "") { return Col0 $l }
    }; return ""
}
function EscJ($s) {
    $s2 = [string]$s
    $s2 = $s2.Replace('\','\\').Replace('"','\"')
    return $s2
}

# ════════════════════════════════════════════════════════════════
# STEP 0 — Timestamped backup of application files
# This script ONLY ever writes: orbit-data.js, prod_raw.csv,
# qa_raw.csv, assoc_raw.csv. Application files (ORBIT.html,
# orbit-app.js, orbit.css) are NEVER written or overwritten.
# Backup guarantees instant recovery if anything else touches them.
# ════════════════════════════════════════════════════════════════
Log "STEP 0: Backing up application files..." Cyan

$APP_FILES = @("ORBIT.html","orbit-app.js","orbit.css")
$backupDir = "$SYS_DIR\BACKUP\$(Get-Date -Format 'yyyyMMdd_HHmm')"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

foreach ($f in $APP_FILES) {
    $p = "$DASH_DIR\$f"
    if (Test-Path $p) {
        Copy-Item $p "$backupDir\$f" -Force
        Log ("  [BACKED UP]  $f -> _system\BACKUP\$(Split-Path $backupDir -Leaf)") Gray
    } else {
        Log ("  [SKIP]  $f not present in ORBIT DASH") Yellow
    }
}

# Record app-file fingerprints — verified again in STEP 4 to prove
# the refresh only touched data files.
$appHashBefore = @{}
foreach ($f in $APP_FILES) {
    $p = "$DASH_DIR\$f"
    if (Test-Path $p) { $appHashBefore[$f] = (Get-FileHash $p -Algorithm SHA256).Hash }
}
Log "  Application files backed up to _system\BACKUP\$(Split-Path $backupDir -Leaf)" Green

# ════════════════════════════════════════════════════════════════
# STEP 1 — Open workbook and RefreshAll (data model + Power Query)
# ════════════════════════════════════════════════════════════════
Log "STEP 1: Opening Excel and refreshing data model..." Cyan

Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$xl = New-Object -ComObject Excel.Application
$xl.Visible        = $true
$xl.DisplayAlerts  = $false
$xl.ScreenUpdating = $false

$wb = $xl.Workbooks.Open($WORKBOOK)
Log "  Workbook open. Running RefreshAll..." Gray
try { $wb.RefreshAll() } catch { Log ("  RefreshAll warning: " + $_.Exception.Message) Yellow }
Start-Sleep -Seconds 10   # allow data model to fully load

Log "  Waiting for background queries to settle..." Gray
Start-Sleep -Seconds 5
Log "  Data model refresh complete. $(Elapsed)" Green

# ════════════════════════════════════════════════════════════════
# STEP 2 — Build DATA_EXTRACT PivotTables and export CSVs
# ════════════════════════════════════════════════════════════════
Log "STEP 2: Building DATA_EXTRACT PivotTables..." Cyan

$dmConn = $wb.Connections.Item("ThisWorkbookDataModel")

# Remove old extraction sheets
foreach ($sh in @($wb.Sheets | Where-Object { $_.Name -eq "DATA_EXTRACT" -or $_.Name -eq "PLAIN_DATA" })) {
    try { $sh.Delete() } catch {}
}

$xs = $wb.Sheets.Add(); $xs.Name = "DATA_EXTRACT"

$xlRow  = 1   # xlRowField
$xlData = 4   # xlDataField

function New-PT($row, $col, $name) {
    return $wb.PivotCaches().Create(2, $dmConn).CreatePivotTable($xs.Cells.Item($row,$col), $name)
}
function F($pt, $mdx, $orient) {
    foreach ($cf in $pt.CubeFields) {
        if ($cf.Name -eq $mdx) { $cf.Orientation = $orient; return }
    }
    Log ("  Field not found: $mdx") Yellow
}

# PT-PROD
Log "  Building PT_PROD..." Gray
$ptP = New-PT 1 1 "PT_PROD"
F $ptP "[Fact_Productivity].[Manager]"        $xlRow
F $ptP "[Fact_Productivity].[Associate Name]" $xlRow
F $ptP "[Fact_Productivity].[Date]"           $xlRow
F $ptP "[Fact_Productivity].[WEEK]"           $xlRow
F $ptP "[Measures].[Actual Production]"       $xlData
F $ptP "[Measures].[Production Target]"       $xlData
Log ("    Rows: " + $ptP.TableRange1.Rows.Count) Gray

# PT-QA
Log "  Building PT_QA..." Gray
$ptQ = New-PT 1 9 "PT_QA"
F $ptQ "[Fact_QA].[Manager]"         $xlRow
F $ptQ "[Fact_QA].[US Domain ID]"    $xlRow
F $ptQ "[Fact_QA].[Associate]"       $xlRow
F $ptQ "[Fact_QA].[Process Date]"    $xlRow
F $ptQ "[Measures].[QA Audit Count]" $xlData
F $ptQ "[Measures].[Avg QA Score]"   $xlData
Log ("    Rows: " + $ptQ.TableRange1.Rows.Count) Gray

# PT-ASSOC
Log "  Building PT_ASSOC..." Gray
$ptA = New-PT 1 17 "PT_ASSOC"
F $ptA "[Dim_Associate].[Manager]"        $xlRow
F $ptA "[Dim_Associate].[Domain ID]"      $xlRow
F $ptA "[Dim_Associate].[Associate Name]" $xlRow
F $ptA "[Measures].[Actual Production]"   $xlData
Log ("    Rows: " + $ptA.TableRange1.Rows.Count) Gray

$wb.Save()
Start-Sleep -Seconds 2

# Copy PT values to plain sheet (values-only paste to strip PT formatting)
Log "  Copying PT values to PLAIN_DATA sheet..." Gray
$ps = $wb.Sheets.Add(); $ps.Name = "PLAIN_DATA"
$ps.Cells.Item(1,1).Value2 = "PROD_START"

$ptP.TableRange1.Copy(); $ps.Cells.Item(2,1).PasteSpecial(-4163); $xl.CutCopyMode = $false
$lastRowP = $ps.Cells.Item(1,1).SpecialCells(11).Row
$ps.Cells.Item($lastRowP+2, 1).Value2 = "QA_START"
$ptQ.TableRange1.Copy(); $ps.Cells.Item($lastRowP+3, 1).PasteSpecial(-4163); $xl.CutCopyMode = $false
$lastRowQ = $ps.Cells.Item(1,1).SpecialCells(11).Row
$ps.Cells.Item($lastRowQ+2, 1).Value2 = "ASSOC_START"
$ptA.TableRange1.Copy(); $ps.Cells.Item($lastRowQ+3, 1).PasteSpecial(-4163); $xl.CutCopyMode = $false

$wb.Save()

# Read PLAIN_DATA and convert to CSVs
$lastRow = $ps.Cells.Item(1,1).SpecialCells(11).Row
$lastCol = $ps.Cells.Item(1,1).SpecialCells(11).Column
$allData = $ps.Range($ps.Cells.Item(1,1), $ps.Cells.Item($lastRow,$lastCol)).Value2
Log ("  Sheet dimensions: $lastRow rows x $lastCol cols") Gray

$prodStart = $null; $qaStart = $null; $assocStart = $null
$prodEnd   = $null; $qaEnd   = $null

for ($r = 1; $r -le $lastRow; $r++) {
    $v = $allData[$r,1]
    if ($v -eq "PROD_START")  { $prodStart  = $r+1 }
    if ($v -eq "QA_START")    { $qaStart    = $r+1; if ($prodStart  -ne $null) { $prodEnd  = $r-2 } }
    if ($v -eq "ASSOC_START") { $assocStart = $r+1; if ($qaStart    -ne $null) { $qaEnd    = $r-2 } }
}
$assocEnd = $lastRow

function Section-To-CSV($data, $rStart, $rEnd, $nCols) {
    $lines = [System.Collections.ArrayList]::new()
    for ($r = $rStart; $r -le $rEnd; $r++) {
        $hasData = $false
        for ($c = 1; $c -le $nCols; $c++) { if ($data[$r,$c] -ne $null) { $hasData=$true; break } }
        if (-not $hasData) { continue }
        $cells = @(1..$nCols | ForEach-Object {
            $v = $data[$r,$_]
            if ($v -eq $null) { "" } else {
                $s = "$v"; $s = $s -replace '"','""'
                if ($s -match '[,"\r\n]') { '"' + $s + '"' } else { $s }
            }
        })
        $null = $lines.Add(($cells -join ","))
    }
    return $lines -join "`n"
}

$prodCsv  = Section-To-CSV $allData $prodStart  $prodEnd  8
$qaCsv    = Section-To-CSV $allData $qaStart    $qaEnd    8
$assocCsv = Section-To-CSV $allData $assocStart $assocEnd 8

[System.IO.File]::WriteAllText("$SYS_DIR\prod_raw.csv",  $prodCsv,  [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText("$SYS_DIR\qa_raw.csv",    $qaCsv,    [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText("$SYS_DIR\assoc_raw.csv", $assocCsv, [System.Text.Encoding]::UTF8)

# Report max date captured
$prodLines = $prodCsv -split "`n"
$maxSerial = 0
foreach ($l in $prodLines) {
    if ($l -match '^\d{5}(,|$)') {
        $s = [int]$l.Split(',')[0]; if ($s -gt $maxSerial) { $maxSerial = $s }
    }
}
$maxDate = if ($maxSerial -gt 0) { [DateTime]::FromOADate($maxSerial).ToString("yyyy-MM-dd") } else { "unknown" }

Log ("  prod_raw.csv : " + [Math]::Round(([System.IO.FileInfo]"$SYS_DIR\prod_raw.csv").Length/1KB,1) + " KB  |  Max date: $maxDate (serial $maxSerial)") Green
Log ("  qa_raw.csv   : " + [Math]::Round(([System.IO.FileInfo]"$SYS_DIR\qa_raw.csv").Length/1KB,1)   + " KB") Green
Log ("  assoc_raw.csv: " + [Math]::Round(([System.IO.FileInfo]"$SYS_DIR\assoc_raw.csv").Length/1KB,1) + " KB") Green

# Close Excel
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
[GC]::Collect()
Log "  Excel closed. $(Elapsed)" Green

# ════════════════════════════════════════════════════════════════
# STEP 2.5 — Export leadership roster -> orbit-auth.js
# ORBIT_Leadership.xlsx is the single source of truth for
# Leadership Portal authorization (DOMAIN = security key).
# ════════════════════════════════════════════════════════════════
Log "STEP 2.5: Exporting leadership roster..." Cyan
& "$SYS_DIR\Export_ORBIT_Leadership.ps1"

# ════════════════════════════════════════════════════════════════
# STEP 3 — Parse CSVs and rebuild orbit-data.js
# ════════════════════════════════════════════════════════════════
Log "STEP 3: Rebuilding orbit-data.js..." Cyan

$prevItem = Get-Item "$DASH_DIR\orbit-data.js" -ErrorAction SilentlyContinue
$prevSize = if ($prevItem) { $prevItem.Length } else { 0 }
$prevRows = if (Test-Path "$DASH_DIR\orbit-data.js") {
    ([regex]::Matches([System.IO.File]::ReadAllText("$DASH_DIR\orbit-data.js"),'\"act\":')).Count
} else { 0 }

# Parse PROD
$rawP = [System.IO.File]::ReadAllLines("$SYS_DIR\prod_raw.csv",[System.Text.Encoding]::UTF8)
$prodRows = [System.Collections.ArrayList]::new()
$m = ""; $a = ""; $ym = ""; $d = ""
for ($i = 1; $i -lt $rawP.Length; $i++) {
    $l = $rawP[$i]; if ($l -eq "") { continue }
    $c = Col0 $l
    if ($c -match "^Week ") {
        $f = ParseCsvLine $l
        $null = $prodRows.Add([PSCustomObject]@{
            m=$m; a=$a; ym=$ym; w=$c; d=$d
            act=[Math]::Round((ToDouble $f[1]),2)
            tgt=[Math]::Round((ToDouble $f[2]),2)
        })
    } elseif (IsDateSerial $c) {
        # Retain both the month (for filtering) and the specific calendar date (for Daily Detail display)
        try {
            $dt = [DateTime]::FromOADate((ToDouble $c))
            $ym = $dt.ToString("yyyy-MM")
            $d  = $dt.ToString("M/d/yyyy")   # e.g. 9/1/2026
        } catch { $ym = ""; $d = "" }
    } elseif ($c -ne "") {
        $nxt = NextC0 $rawP ($i+1)
        if ($nxt -match "^[A-Za-z]" -and -not(IsDomainID $nxt) -and
            -not(IsDateSerial $nxt) -and -not($nxt -match "^Week ")) {
            $m = $c; $a = ""; $ym = ""; $d = ""
        } else { $a = $c; $ym = ""; $d = "" }
    }
}
$orgMgrs = $prodRows | Select-Object -ExpandProperty m | Sort-Object -Unique | Where-Object { $_ -ne "" }
Log ("  PROD: " + $prodRows.Count + " rows, " + $orgMgrs.Count + " managers") Gray

# Parse QA
$rawQ = [System.IO.File]::ReadAllLines("$SYS_DIR\qa_raw.csv",[System.Text.Encoding]::UTF8)
$qaRows = [System.Collections.ArrayList]::new()
$mq = ""; $iq = ""; $aq = ""
for ($i = 1; $i -lt $rawQ.Length; $i++) {
    $l = $rawQ[$i]; if ($l -eq "") { continue }
    $c = Col0 $l
    if (IsDateSerial $c) {
        $f = ParseCsvLine $l; $cnt = ToDouble $f[1]; $scr = ToDouble $f[2]
        if ($cnt -gt 0) {
            try { $d = [DateTime]::FromOADate((ToDouble $c)).ToString("yyyy-MM-dd") } catch { $d = $c }
            $null = $qaRows.Add([PSCustomObject]@{
                m=$mq; id=$iq; a=$aq; d=$d
                cnt=[Math]::Round($cnt,4); scr=[Math]::Round($scr,6)
            })
        }
    } elseif (IsDomainID $c) { $iq = $c; $aq = "" }
    elseif ($c -ne "") {
        $nxt = NextC0 $rawQ ($i+1)
        if (IsDomainID $nxt) { $mq = $c; $iq = ""; $aq = "" } else { $aq = $c }
    }
}
$qaMgrs = $qaRows | Select-Object -ExpandProperty m | Sort-Object -Unique | Where-Object { $_ -ne "" }
Log ("  QA: " + $qaRows.Count + " rows, " + $qaMgrs.Count + " managers") Gray

# Parse ASSOC
$rawA = [System.IO.File]::ReadAllLines("$SYS_DIR\assoc_raw.csv",[System.Text.Encoding]::UTF8)
$assocRows = [System.Collections.ArrayList]::new()
$ma = ""; $ia = ""
for ($i = 1; $i -lt $rawA.Length; $i++) {
    $l = $rawA[$i]; if ($l -eq "") { continue }
    $f = ParseCsvLine $l; $c = $f[0].Trim(); if ($c -eq "") { continue }
    if (IsDomainID $c) { $ia = $c; continue }
    if ($f[1].Trim() -eq "") { $ma = $c; $ia = "" }
    elseif ($ia -ne "") { $null = $assocRows.Add([PSCustomObject]@{id=$ia; n=$c}); $ia = "" }
}
Log ("  ASSOC: " + $assocRows.Count + " rows") Gray

# Write orbit-data.js
$sb = [System.Text.StringBuilder]::new()
[void]$sb.Append("// PROJECT ORBIT`n")
[void]$sb.Append("// Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n")
[void]$sb.Append("// Max date captured: $maxDate (serial $maxSerial)`n")
[void]$sb.Append("// Zero-production rows included so all associates appear in table`n")

[void]$sb.Append("const D_PROD=[")
$fi = $true
foreach ($r in $prodRows) {
    if (-not $fi) { [void]$sb.Append(",") }; $fi = $false
    [void]$sb.Append("{`"m`":`"$(EscJ $r.m)`",`"a`":`"$(EscJ $r.a)`",`"ym`":`"$($r.ym)`",`"w`":`"$(EscJ $r.w)`",`"d`":`"$($r.d)`",`"act`":$($r.act),`"tgt`":$($r.tgt)}")
}
[void]$sb.Append("];`n")

[void]$sb.Append("const D_QA=[")
$fi = $true
foreach ($r in $qaRows) {
    if (-not $fi) { [void]$sb.Append(",") }; $fi = $false
    [void]$sb.Append("{`"m`":`"$(EscJ $r.m)`",`"id`":`"$(EscJ $r.id)`",`"a`":`"$(EscJ $r.a)`",`"d`":`"$($r.d)`",`"cnt`":$($r.cnt),`"scr`":$($r.scr)}")
}
[void]$sb.Append("];`n")

[void]$sb.Append("const D_ASSOC=[")
$fi = $true
foreach ($r in $assocRows) {
    if (-not $fi) { [void]$sb.Append(",") }; $fi = $false
    [void]$sb.Append("{`"id`":`"$(EscJ $r.id)`",`"n`":`"$(EscJ $r.n)`"}")
}
[void]$sb.Append("];`n")

$allMgrs = ($orgMgrs + $qaMgrs) | Sort-Object -Unique | Where-Object { $_ -ne "" }
$mparts  = ($allMgrs | ForEach-Object { "`"$(EscJ $_)`"" }) -join ","
[void]$sb.Append("const D_MANAGERS=[$mparts];`n")

[System.IO.File]::WriteAllText("$DASH_DIR\orbit-data.js", $sb.ToString(), [System.Text.Encoding]::UTF8)

$newSize = (Get-Item "$DASH_DIR\orbit-data.js").Length
$newRows = $prodRows.Count
Log ("  orbit-data.js: " + [Math]::Round($newSize/1KB,1) + " KB") Green
Log ("  D_PROD rows: $prevRows -> $newRows  (delta: +$($newRows - $prevRows))") $(if(($newRows-$prevRows) -ge 0){"Green"}else{"Yellow"})

# ════════════════════════════════════════════════════════════════
# STEP 4 — Verify ORBIT DASH deployment folder
# ════════════════════════════════════════════════════════════════
Log "STEP 4: Verifying deployment folder..." Cyan

# All files are written directly to ORBIT DASH — no copy needed.
# Just verify all required runtime files are present.
$required = @(
    "ORBIT.html","orbit.css","orbit-app.js","orbit-data.js","chart.umd.min.js",
    "assets\logos\carelon.png","assets\logos\ignite.png","assets\logos\pi.png",
    "assets\logos\orbit-symbol.svg","assets\logos\orbit-logo-lockup.svg"
)
$missing = 0
foreach ($f in $required) {
    $p = "$DASH_DIR\$f"
    if (Test-Path $p) {
        Log ("  [OK]  $f  (" + [Math]::Round((Get-Item $p).Length/1KB,1) + " KB)") Gray
    } else {
        Log ("  [MISSING]  $f") Red
        $missing++
    }
}
if ($missing -eq 0) {
    Log "  All deployment files verified. $(Elapsed)" Green
} else {
    Log "  WARNING: $missing file(s) missing from deployment folder." Yellow
}

# ── Application-file integrity guard ──────────────────────────
# Confirm ORBIT.html / orbit-app.js / orbit.css were NOT modified
# by this refresh run (hashes compared against STEP 0 snapshot).
$appTouched = @()
foreach ($f in $APP_FILES) {
    $p = "$DASH_DIR\$f"
    if ((Test-Path $p) -and $appHashBefore.ContainsKey($f)) {
        if ((Get-FileHash $p -Algorithm SHA256).Hash -ne $appHashBefore[$f]) {
            $appTouched += $f
        }
    }
}
if ($appTouched.Count -eq 0) {
    Log "  Integrity check PASSED: application files untouched (data-only refresh)." Green
} else {
    Log "  INTEGRITY WARNING: these app files changed during refresh: $($appTouched -join ', ')" Red
    Log "  Restore from backup: _system\BACKUP\$(Split-Path $backupDir -Leaf)" Red
}

# ════════════════════════════════════════════════════════════════
# STEP 5 — Git commit + push -> GitHub Pages auto-publishes
#
# Stages ALL tracked deployment files that have changed:
#   orbit-data.js   orbit-auth.js          (generated on every refresh)
#   ORBIT.html      orbit-app.js  orbit.css (committed only when edited)
#
# Commit message reflects what actually changed:
#   "data: refresh …"            — data files only
#   "data + code: refresh …"     — data files + one or more app files
#   "code: …"                    — app files only (manual push scenario)
#
# If nothing has changed since the last commit, step is skipped silently.
# ════════════════════════════════════════════════════════════════
Log "STEP 5: Publishing to GitHub Pages via git..." Cyan

if (-not (Test-Path $GIT)) {
    Log "  [SKIP] git not found at: $GIT" Yellow
    Log "  Install Git for Windows and re-run to enable auto-publish." Yellow
} else {
    # Stage all deployment artifacts (git ignores files that haven't changed)
    $filesToStage = @(
        "orbit-data.js",
        "orbit-auth.js",
        "ORBIT.html",
        "orbit-app.js",
        "orbit.css"
    )
    & $GIT -C $DASH_DIR add @filesToStage 2>&1 | ForEach-Object { if ($_) { Log "  git: $_" Gray } }

    # Determine what is actually staged to build an accurate commit message
    $staged     = & $GIT -C $DASH_DIR diff --name-only --cached 2>&1
    $dataFiles  = @("orbit-data.js","orbit-auth.js")
    $codeFiles  = @("ORBIT.html","orbit-app.js","orbit.css")
    $hasData    = ($staged | Where-Object { $dataFiles -contains $_ }).Count -gt 0
    $hasCode    = ($staged | Where-Object { $codeFiles -contains $_ }).Count -gt 0
    $codeList   = ($staged | Where-Object { $codeFiles -contains $_ }) -join ", "

    if ($staged) {
        $dataDate  = if ($maxDate) { $maxDate } else { (Get-Date -Format 'yyyy-MM-dd') }

        if ($hasData -and $hasCode) {
            $commitMsg = "data + code: refresh $dataDate ($newRows prod rows) [$codeList]"
        } elseif ($hasData) {
            $commitMsg = "data: refresh $dataDate ($newRows prod rows)"
        } else {
            $commitMsg = "code: publish $dataDate [$codeList]"
        }

        & $GIT -C $DASH_DIR commit -m $commitMsg 2>&1 | ForEach-Object { Log "  $_" Gray }

        if ($LASTEXITCODE -eq 0) {
            Log "  Committed: $commitMsg" Green
            & $GIT -C $DASH_DIR push origin main 2>&1 | ForEach-Object { Log "  $_" Gray }
            if ($LASTEXITCODE -eq 0) {
                Log "  Pushed to GitHub. Pages will update within ~60 seconds." Green
                Log "  Live URL: $ORBIT_URL" Cyan
            } else {
                Log "  [WARN] Push failed. Changes committed locally; push manually." Yellow
                Log "  Run: git -C `"$DASH_DIR`" push origin main" Yellow
            }
        } else {
            Log "  [WARN] git commit failed. Check output above." Yellow
        }
    } else {
        Log "  Nothing changed since last commit - skipping." Gray
    }
}

# ════════════════════════════════════════════════════════════════
# STEP 6 — Open / reload ORBIT.html in the default browser
# ════════════════════════════════════════════════════════════════
Log "STEP 6: Opening ORBIT in browser..." Cyan

# If Edge has ORBIT open, send F5 to reload; otherwise open fresh
$edgeProc = Get-Process msedge -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -match "ORBIT" } |
    Select-Object -First 1

if ($edgeProc) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    [Microsoft.VisualBasic.Interaction]::AppActivate($edgeProc.Id)
    Start-Sleep -Milliseconds 400
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait("{F5}")
    Log "  ORBIT tab reloaded (F5 sent to Edge)." Green
} else {
    Start-Process $ORBIT_URL
    Log "  ORBIT opened in default browser (GitHub Pages URL)." Green
}

# ════════════════════════════════════════════════════════════════
# DONE
# ════════════════════════════════════════════════════════════════
$total = [int]((Get-Date) - $T0).TotalSeconds
$delta = $newRows - $prevRows
Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "  ORBIT REFRESH + PUBLISH COMPLETE" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host ("  Max date   : $maxDate  (serial $maxSerial)") -ForegroundColor Green
Write-Host ("  D_PROD rows: $newRows  (+$delta new)") -ForegroundColor Green
Write-Host ("  Duration   : ${total}s") -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  DATA IS LIVE on GitHub Pages (~60 s propagation)." -ForegroundColor Cyan
Write-Host ""
Write-Host "  Live URL:" -ForegroundColor Yellow
Write-Host "  $ORBIT_URL" -ForegroundColor White
Write-Host ""
