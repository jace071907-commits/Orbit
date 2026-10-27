# ============================================================
# PROJECT ORBIT - PowerShell Pre-Publish Validator
#
# Runs BEFORE any Refresh_ORBIT.ps1 (or other .ps1) change is
# committed/pushed. Two gates, both must pass:
#
#   GATE 1 - Parser: [System.Management.Automation.Language.Parser]
#            ::ParseFile() must report 0 errors.
#
#   GATE 2 - String hygiene: token-level scan. Any StringLiteral,
#            StringExpandable, or HereString token containing a
#            non-ASCII character (> U+007F) FAILS the file.
#            Comments are NOT scanned (Unicode is allowed there).
#
# Why: powershell.exe (PS 5.1) decodes BOM-less UTF-8 as ANSI.
#      An em-dash inside a string decodes to bytes that include a
#      curly-quote (U+201D), which the tokenizer treats as a string
#      delimiter - terminating strings early and producing cascading
#      "Missing closing '}'" errors. (Incident 2026-10-19.)
#
# Usage:
#   powershell -ExecutionPolicy Bypass -NoProfile -File Validate_ORBIT.ps1
#   powershell ... -File Validate_ORBIT.ps1 -Path <file-or-dir>
#
# Exit code: 0 = all files clean, 1 = one or more failures.
# ============================================================
param(
    [string]$Path = $PSScriptRoot   # default: validate every .ps1 in _system\
)

$script:failures = 0

function Log($msg, $color="White") {
    $ts = (Get-Date).ToString("HH:mm:ss")
    Write-Host "[$ts] $msg" -ForegroundColor $color
}

# ── String-token kinds scanned for non-ASCII ────────────────────────────
# Comments, comment blocks, and XML doc comments are excluded by design.
$STRING_KINDS = @(
    'StringLiteral',        # 'single quoted'
    'StringExpandable',     # "double quoted"
    'HereStringLiteral',    # @'...'@
    'HereStringExpandable'  # @"..."@
)

function Test-OneFile($file) {
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $file.FullName, [ref]$tokens, [ref]$errors
    )

    $fileFailed = $false

    # ── GATE 1: parser errors ──
    if ($errors.Count -gt 0) {
        $fileFailed = $true
        Log ("  [FAIL] {0} - {1} parser error(s)" -f $file.Name, $errors.Count) Red
        $errors | Select-Object -First 10 | ForEach-Object {
            Log ("         {0}:{1}  {2}" -f `
                $_.Extent.StartLineNumber, $_.Extent.StartColumnNumber, $_.Message) Red
        }
        if ($errors.Count -gt 10) {
            Log ("         ... and {0} more" -f ($errors.Count - 10)) Red
        }
    }

    # ── GATE 2: non-ASCII inside string tokens ──
    foreach ($t in $tokens) {
        if ($STRING_KINDS -contains $t.Kind) {
            $bad = $t.Text.ToCharArray() | Where-Object { [int]$_ -gt 127 }
            if ($bad) {
                $fileFailed = $true
                $list = ($bad | ForEach-Object { 'U+{0:X4}' -f [int]$_ } |
                         Select-Object -Unique) -join ' '
                Log ("  [FAIL] {0} line {1}: non-ASCII in string ({2})" -f `
                    $file.Name, $t.Extent.StartLineNumber, $list) Red
                $snippet = $t.Text
                if ($snippet.Length -gt 70) { $snippet = $snippet.Substring(0,70) + "..." }
                Log ("         > $snippet") DarkRed
            }
        }
    }

    if (-not $fileFailed) {
        Log ("  [PASS] {0}" -f $file.Name) Green
    }
    return $fileFailed
}

# ── Resolve targets ──
if (Test-Path $Path -PathType Leaf) {
    $targets = @(Get-Item $Path)
} else {
    $targets = Get-ChildItem $Path -Filter "*.ps1" -File
}

if ($targets.Count -eq 0) {
    Log "No .ps1 files found at: $Path" Yellow
    exit 0
}

Log "Validating $($targets.Count) PowerShell file(s)..." Cyan
foreach ($f in $targets) {
    if (Test-OneFile $f) { $script:failures++ }
}

Write-Host ""
if ($script:failures -eq 0) {
    Log "VALIDATION PASSED - $($targets.Count) file(s), 0 parser errors, 0 non-ASCII strings." Green
    exit 0
} else {
    Log "VALIDATION FAILED - $($script:failures) of $($targets.Count) file(s) did not pass." Red
    Log "Do NOT commit or push until all files pass." Red
    exit 1
}
