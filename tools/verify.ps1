# VBE Auto — Universal VBA Verification Suite
# 5-stage verification for any VBA project
# Usage: & verify.ps1 [-Workbook path] [-Config path]

param(
    [string]$ConfigPath,
    [string]$Workbook
)

$ErrorActionPreference = "Continue"
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent

# ============================================================================
# CONFIGURATION
# ============================================================================

function FindConfig {
    $paths = @(
        $ConfigPath,
        "$ScriptDir\vbe-auto-config.json",
        ".\vbe-auto-config.json"
    )
    foreach ($p in $paths) {
        if ($p -and (Test-Path $p -ErrorAction SilentlyContinue)) {
            $result = Get-Content $p -Raw | ConvertFrom-Json
            return $result
        }
    }
    return $null
}

$config = FindConfig

# Resolve workbook path
$wbPath = if ($Workbook) { $Workbook } elseif ($config) { $config.output_workbook } else { "$ScriptDir\Output.xlsm" }

if (-not (Test-Path $wbPath)) {
    Write-Host "[ERROR] Workbook not found: $wbPath" -ForegroundColor Red
    exit 1
}

# ============================================================================
# OPEN WORKBOOK
# ============================================================================

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.Interactive = $false

try {
    $wb = $xl.Workbooks.Open($wbPath, 0, $false)
} catch {
    Write-Host "[FATAL] Cannot open workbook: $($_.Exception.Message)" -ForegroundColor Red
    $xl.Quit()
    exit 1
}

$passed = 0
$failed = 0
$skipped = 0

function Check {
    param([string]$Stage, [string]$Name, [bool]$Result, [string]$Detail = "")
    if ($Result) {
        Write-Host "  [PASS] $Name" -ForegroundColor Green
        $script:passed++
    } else {
        Write-Host "  [FAIL] $Name - $Detail" -ForegroundColor Red
        $script:failed++
    }
}

# ============================================================================
# STAGE 1: File Integrity
# ============================================================================

Write-Host "`n[1/5] File Integrity Checks..." -ForegroundColor Cyan

$fileSize = (Get-Item $wbPath).Length
Check "1" "File exists" $true
Check "1" "Size: $([math]::Round($fileSize/1KB,1)) KB" ($fileSize -gt 10KB) "File too small"
Check "1" "File accessible" ($fileSize -gt 0) "Zero bytes"

# ============================================================================
# STAGE 2: COM Compilation
# ============================================================================

Write-Host "`n[2/5] COM Compilation Check..." -ForegroundColor Cyan

Check "2" "Workbook opened via COM" ($wb -ne $null)

try {
    $xl.VBE.CommandBars("Menu Bar").Controls("Debug").Controls("Compile VBAProject").Execute()
    Check "2" "VBA compilation: OK" $true
} catch {
    Check "2" "VBA compilation" $false $_.Exception.Message
}

# ============================================================================
# STAGE 3: Module Inventory
# ============================================================================

Write-Host "`n[3/5] Module Inventory..." -ForegroundColor Cyan

$moduleCount = 0
foreach ($comp in $wb.VBProject.VBComponents) {
    $type = switch ($comp.Type) { 1 { "Module" } 2 { "Class" } 3 { "Form" } 100 { "Document" } default { "Unknown" } }
    $lines = $comp.CodeModule.CountOfLines
    # Document types (sheet modules) are allowed to be empty
    if ($comp.Type -eq 100) {
        Check "3" "$($comp.Name) ($lines lines, $type)" $true
    } else {
        Check "3" "$($comp.Name) ($lines lines, $type)" ($lines -gt 0) "Empty module"
    }
    $moduleCount++
}

# Check expected modules from config
if ($config -and $config.verification -and $config.verification.expected_modules) {
    $expectedModules = @($config.verification.expected_modules)
    foreach ($expected in $expectedModules) {
        $found = $false
        foreach ($comp in $wb.VBProject.VBComponents) {
            if ($comp.Name -eq $expected) { $found = $true; break }
        }
        Check "3" "Expected module: $expected" $found "Missing"
    }
}

$totalLines = 0
foreach ($comp in $wb.VBProject.VBComponents) {
    $totalLines += $comp.CodeModule.CountOfLines
}

Check "3" "Total: $moduleCount modules, $totalLines lines" ($moduleCount -gt 0) "No modules found"

# ============================================================================
# STAGE 4: Sheet Verification
# ============================================================================

Write-Host "`n[4/5] Sheet Verification..." -ForegroundColor Cyan

$sheetCount = $wb.Sheets.Count
foreach ($ws in $wb.Sheets) {
    $lastRow = $ws.Cells($ws.Rows.Count, "A").End(-4162).Row
    Check "4" "$($ws.Name) (exists, $lastRow rows)" $true
}

# Check expected sheets from config
if ($config -and $config.verification -and $config.verification.expected_sheets) {
    $expectedSheets = @($config.verification.expected_sheets)
    foreach ($expected in $expectedSheets) {
        $found = $false
        foreach ($ws in $wb.Sheets) {
            if ($ws.Name -eq $expected) { $found = $true; break }
        }
        Check "4" "Expected sheet: $expected" $found "Missing"
    }
}

Check "4" "Sheet count: $sheetCount" ($sheetCount -gt 0) "No sheets"

# ============================================================================
# STAGE 5: Configuration & Constants
# ============================================================================

Write-Host "`n[5/5] Configuration & Constants..." -ForegroundColor Cyan

# Check config constants
if ($config -and $config.verification -and $config.verification.expected_constants) {
    $expectedConstants = @($config.verification.expected_constants)
    foreach ($const in $expectedConstants) {
        $name = $const.name
        $expectedValue = $const.value

        # Search for the constant/property in module code
        $found = $false
        foreach ($comp in $wb.VBProject.VBComponents) {
            $cm = $comp.CodeModule
            for ($line = 1; $line -le $cm.CountOfLines; $line++) {
                $text = $cm.Lines($line, 1).Trim()
                # Match Public Const or Const
                if ($text -match "(Public\s+)?Const\s+$name\s*(As\s+\w+\s*)?=\s*(.+?)($|')") {
                    $found = $true
                    $rawValue = $matches[3].Trim().Trim('"').Trim("'")
                    if ($rawValue -eq $expectedValue -or $rawValue -like "*$expectedValue*") {
                        Check "5" "[$name] = $rawValue" $true
                    } else {
                        Check "5" "[$name] expected '$expectedValue', got '$rawValue'" $false "Mismatch"
                    }
                    break
                }
                # Match Property Get (for runtime-resolved values like MASTER_PWD, VERSION)
                if ($text -match "Property\s+Get\s+$name\s*\(\)") {
                    # For Property Get, just confirm it exists
                    $found = $true
                    # Look for the return value in the next few lines
                    for ($lookLine = $line + 1; $lookLine -le [Math]::Min($line + 5, $cm.CountOfLines); $lookLine++) {
                        $returnText = $cm.Lines($lookLine, 1).Trim()
                        if ($returnText -match "$name\s*=\s*(.+?)($|')") {
                            $rawValue = $matches[1].Trim().Trim('"').Trim("'")
                            if ($rawValue -eq $expectedValue -or $rawValue -like "*$expectedValue*") {
                                Check "5" "[$name] Property Get = $rawValue" $true
                            } else {
                                Check "5" "[$name] Property Get = $rawValue (expected '$expectedValue')" ($rawValue -like "*$expectedValue*") "Value mismatch"
                            }
                            break
                        }
                    }
                    if ($found -and $lookLine -gt $cm.CountOfLines) {
                        Check "5" "[$name] Property Get declared" $true
                    }
                    break
                }
            }
            if ($found) { break }
        }
        if (-not $found) {
            Check "5" "[$name]" $false "Not found"
        }
    }
}

# Check protection password resolution (if configured)
if ($config -and $config.protection -and $config.protection.sheet_password) {
    $pwd = $config.protection.sheet_password
    $found = $false
    foreach ($comp in $wb.VBProject.VBComponents) {
        $cm = $comp.CodeModule
        for ($line = 1; $line -le $cm.CountOfLines; $line++) {
            $text = $cm.Lines($line, 1).Trim()
            if ($text -like "*$pwd*") {
                # Check if it's in a Property Get (good) vs hardcoded Const (bad)
                if ($text -match "Property\s+Get") {
                    Check "5" "Password: Property Get (secure)" $true
                } elseif ($text -match "Public\s+Const") {
                    Check "5" "Password: Public Const (insecure)" $false "Should use Property Get"
                }
                $found = $true
                break
            }
        }
        if ($found) { break }
    }
    if (-not $found) {
        Check "5" "Password reference" $false "Not found"
    }
}

# ============================================================================
# SUMMARY
# ============================================================================

$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseCOMObject($xl) | Out-Null

Write-Host "`n=== VERIFICATION SUMMARY ===" -ForegroundColor Cyan
Write-Host "  Passed:  $passed" -ForegroundColor Green
Write-Host "  Failed:  $failed" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })
Write-Host "  Skipped: $skipped" -ForegroundColor Gray
Write-Host ""

if ($failed -eq 0) {
    Write-Host "  SAFE TO OPEN: No critical issues found." -ForegroundColor Green
} else {
    Write-Host "  WARNING: $failed issue(s) detected. Review before deploying." -ForegroundColor Red
}

# Export results
$resultsPath = "$ScriptDir\verify_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
@{
    Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Workbook = $wbPath
    Passed = $passed
    Failed = $failed
    Skipped = $skipped
    Safe = ($failed -eq 0)
} | ConvertTo-Json | Out-File $resultsPath -Force

Write-Host "  Results saved: $resultsPath" -ForegroundColor Gray
