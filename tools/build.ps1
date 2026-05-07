# VBE Auto — Universal VBA Build Script
# Build ANY VBA project from source files (.bas, .frm, .cls)
# Usage: & build.ps1 [-Config path] [-Master path] [-Source path] [-Output path]

param(
    [string]$ConfigPath,
    [string]$Master,
    [string]$Source,
    [string]$Output
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent

# ============================================================================
# CONFIGURATION
# ============================================================================

# Resolve config
$config = $null
$configPaths = @(
    $ConfigPath,
    "$ScriptDir\vbe-auto-config.json",
    ".\vbe-auto-config.json",
    "$PSScriptRoot\vbe-auto-config.json"
)
foreach ($cp in $configPaths) {
    if ($cp -and (Test-Path $cp -ErrorAction SilentlyContinue)) {
        $configRaw = Get-Content $cp -Raw
        $config = $configRaw | ConvertFrom-Json
        break
    }
}

# Resolve paths: CLI arg > Config > Default
$masterPath = if ($Master) { $Master } elseif ($config) { $config.master_workbook } else { "$ScriptDir\Master.xlsm" }
$sourceDir = if ($Source) { $Source } elseif ($config) { $config.vba_source_dir } else { "$ScriptDir\VBA_Modules" }
$outputPath = if ($Output) { $Output } elseif ($config) { $config.output_workbook } else { "$ScriptDir\Output.xlsm" }
$thisWorkbookSource = if ($config -and $config.thisworkbook_source) { $config.thisworkbook_source } else { "$sourceDir\ThisWorkbook.cls" }

Write-Host "`n=== VBE Auto Build ===" -ForegroundColor Cyan
Write-Host "Config : $(if ($config) { 'Loaded' } else { 'None' })"
Write-Host "Master : $masterPath"
Write-Host "Source : $sourceDir"
Write-Host "Output : $outputPath"
Write-Host ""

# Validate paths
if (-not (Test-Path $masterPath)) {
    Write-Host "[ERROR] Master workbook not found: $masterPath" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $sourceDir)) {
    Write-Host "[ERROR] Source directory not found: $sourceDir" -ForegroundColor Red
    exit 1
}

# ============================================================================
# STEP 1: Kill Excel
# ============================================================================

Write-Host "[1/7] Killing Excel..." -ForegroundColor Yellow
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2

# ============================================================================
# STEP 2: Open Master Workbook
# ============================================================================

Write-Host "[2/7] Opening MASTER workbook..." -ForegroundColor Yellow
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$xl.EnableEvents = $false
$xl.ScreenUpdating = $false

try {
    $wb = $xl.Workbooks.Open($masterPath, 0, $false)
    Write-Host "  Sheets: $($wb.Sheets.Count)" -ForegroundColor Gray
} catch {
    Write-Host "[ERROR] Failed to open master: $($_.Exception.Message)" -ForegroundColor Red
    $xl.Quit()
    exit 1
}

# ============================================================================
# STEP 3: Strip All User Modules
# ============================================================================

Write-Host "[3/7] Stripping all user modules..." -ForegroundColor Yellow
$removed = 0
$components = @()

# Collect component names first (can't iterate and modify simultaneously)
foreach ($comp in $wb.VBProject.VBComponents) {
    if ($comp.Type -in @(1, 2, 3)) {  # Module, Class, Form
        $components += $comp.Name
    }
}

foreach ($name in $components) {
    try {
        $wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item($name))
        $removed++
    } catch {
        Write-Host "  Warning: Could not remove $name" -ForegroundColor Gray
    }
}
Write-Host "  Removed $removed modules" -ForegroundColor Gray

# ============================================================================
# STEP 4: Import Source Files
# ============================================================================

Write-Host "[4/7] Importing source files..." -ForegroundColor Yellow

# Import .bas files
$basFiles = Get-ChildItem -Path $sourceDir -Filter "*.bas" -ErrorAction SilentlyContinue
foreach ($f in $basFiles) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    # Skip ThisWorkbook as it needs special handling
    if ($name -eq "ThisWorkbook") { continue }

    try {
        $comp = $wb.VBProject.VBComponents.Import($f.FullName)
        Write-Host "  Imported $name ($($comp.CodeModule.CountOfLines) lines)" -ForegroundColor Gray
    } catch {
        Write-Host "  FAILED: $name - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Import .frm files
$frmFiles = Get-ChildItem -Path $sourceDir -Filter "*.frm" -ErrorAction SilentlyContinue
foreach ($f in $frmFiles) {
    try {
        $comp = $wb.VBProject.VBComponents.Import($f.FullName)
        Write-Host "  Imported $([System.IO.Path]::GetFileNameWithoutExtension($f.Name)) ($($comp.CodeModule.CountOfLines) lines)" -ForegroundColor Gray
    } catch {
        Write-Host "  FAILED: $f - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Inject ThisWorkbook
if (Test-Path $thisWorkbookSource) {
    try {
        $twb = $wb.VBProject.VBComponents.Item("ThisWorkbook")
        $twb.CodeModule.DeleteLines(1, $twb.CodeModule.CountOfLines)
        $rawCode = Get-Content $thisWorkbookSource -Raw
        # Strip attribute lines (document modules already have these internally)
        $code = ($rawCode -split "`n" | Where-Object { $_ -notmatch '^\s*Attribute\s+' }) -join "`n"
        $twb.CodeModule.AddFromString($code)
        Write-Host "  Injected ThisWorkbook" -ForegroundColor Gray
    } catch {
        Write-Host "  FAILED: ThisWorkbook - $($_.Exception.Message)" -ForegroundColor Red
    }
}

$basCount = $basFiles.Count
$frmCount = $frmFiles.Count
Write-Host "  Imported $basCount .bas, $frmCount .frm files" -ForegroundColor Gray

# ============================================================================
# STEP 5: Compile
# ============================================================================

Write-Host "[5/7] Compiling..." -ForegroundColor Yellow
try {
    $xl.VBE.CommandBars("Menu Bar").Controls("Debug").Controls("Compile VBAProject").Execute()
    Write-Host "  COMPILE: OK" -ForegroundColor Green
} catch {
    Write-Host "  COMPILE: FAILED - $($_.Exception.Message)" -ForegroundColor Red
    $wb.Close($false)
    $xl.Quit()
    exit 1
}

# ============================================================================
# STEP 6: Save
# ============================================================================

Write-Host "[6/7] Saving output..." -ForegroundColor Yellow

# Ensure output directory exists
$outDir = Split-Path $outputPath -Parent
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

# Save as new file
try {
    # Delete existing output if it exists
    if (Test-Path $outputPath) {
        Remove-Item $outputPath -Force
    }
    $wb.SaveAs($outputPath, 52)  # xlOpenXMLWorkbookMacroEnabled
    $size = (Get-Item $outputPath).Length
    Write-Host "  Saved: $([math]::Round($size/1KB,1)) KB" -ForegroundColor Gray
} catch {
    Write-Host "  SAVE FAILED: $($_.Exception.Message)" -ForegroundColor Red
    $wb.Close($false)
    $xl.Quit()
    exit 1
}

# ============================================================================
# STEP 7: Cleanup
# ============================================================================

Write-Host "[7/7] Cleaning up..." -ForegroundColor Yellow
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseCOMObject($xl) | Out-Null

Write-Host "`n=== BUILD COMPLETE ===" -ForegroundColor Green
Write-Host "File: $(Split-Path $outputPath -Leaf)" -ForegroundColor Cyan
Write-Host "Size: $([math]::Round((Get-Item $outputPath).Length/1KB,1)) KB" -ForegroundColor Cyan
Write-Host "Modified: $((Get-Item $outputPath).LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
Write-Host "Path: $outputPath" -ForegroundColor Gray
