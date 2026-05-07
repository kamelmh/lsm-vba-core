# VBE Auto — Universal Excel VBE Control Suite
# Project-agnostic VBE session management, macro execution, module control
# Usage: . .\vbe.ps1   (dot-source to load all functions)
# Commands: vbe open, vbe close, vbe save, vbe compile, vbe macro, vbe macro-all, etc.

$ErrorActionPreference = "Continue"

# ============================================================================
# CONFIGURATION — Auto-discover or use defaults
# ============================================================================

# Auto-discover config
$VBEConfig = $null
$VBEConfigPaths = @(
    "$PSScriptRoot\vbe-auto-config.json",
    "$env:USERPROFILE\Desktop\vbe-auto\config.json",
    ".\vbe-auto-config.json",
    "$PSScriptRoot\..\config.json"
)
foreach ($_vbeCfgPath in $VBEConfigPaths) {
    if ($_vbeCfgPath -and (Test-Path $_vbeCfgPath -ErrorAction SilentlyContinue)) {
        $VBEConfig = Get-Content $_vbeCfgPath -Raw | ConvertFrom-Json
        break
    }
}
$VBEToolkitDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$VBELogPath = "$VBEToolkitDir\vbe-auto.log"
$VBEStatePath = "$VBEToolkitDir\vbe-auto.state.json"

# Session object (global)
$VBE = [PSCustomObject]@{
    Excel = $null
    Workbook = $null
    Config = $VBEConfig
    SessionStart = $null
    CommandCount = 0
    LastError = $null
    TransactionLog = @()
}

# ============================================================================
# LOGGING
# ============================================================================

function vbe-log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $VBELogPath -Value $logEntry -Force
    $VBE.TransactionLog += $logEntry
    return $logEntry
}

# ============================================================================
# STATE MANAGEMENT
# ============================================================================

function vbe-state-save {
    $state = [PSCustomObject]@{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Project = if ($VBE.Config) { $VBE.Config.project_name } else { "Unknown" }
        Workbook = if ($VBE.Workbook) { $VBE.Workbook.Name } else { "None" }
        FilePath = if ($VBE.Workbook) { $VBE.Workbook.FullName } else { "N/A" }
        Sheets = if ($VBE.Workbook) { $VBE.Workbook.Sheets.Count } else { 0 }
        Modules = if ($VBE.Workbook) { $VBE.Workbook.VBProject.VBComponents.Count } else { 0 }
        Commands = $VBE.CommandCount
        SessionStart = if ($VBE.SessionStart) { $VBE.SessionStart.ToString('HH:mm:ss') } else { "N/A" }
        Duration = if ($VBE.SessionStart) { "$([math]::Round((Get-Date - $VBE.SessionStart).TotalMinutes, 1)) min" } else { "N/A" }
    }
    $state | ConvertTo-Json -Depth 3 | Out-File $VBEStatePath -Force
}

function vbe-state-load {
    if (Test-Path $VBEStatePath) {
        return Get-Content $VBEStatePath | ConvertFrom-Json
    }
    return $null
}

# ============================================================================
# SESSION MANAGEMENT
# ============================================================================

function vbe-open {
    param(
        [switch]$Visible,
        [switch]$NoEvents,
        [string]$Path,
        [switch]$ListProjects
    )

    if ($ListProjects.IsPresent) {
        Write-Host "`n=== Available Projects ===" -ForegroundColor Cyan
        $configPaths = Get-ChildItem -Path "$env:USERPROFILE\Desktop" -Filter "vbe-auto-config.json" -Recurse -ErrorAction SilentlyContinue
        if ($configPaths) {
            foreach ($cfg in $configPaths) {
                $config = Get-Content $cfg.FullName | ConvertFrom-Json
                $proj = $config.project_name
                $ver = $config.version
                $master = $config.master_workbook
                Write-Host "  $proj ($ver)" -ForegroundColor Yellow
                Write-Host "    Config: $($cfg.FullName)" -ForegroundColor Gray
                Write-Host "    Master: $master" -ForegroundColor Gray
            }
        } else {
            Write-Host "  No projects found. Create a vbe-auto-config.json to register a project." -ForegroundColor Gray
        }
        return
    }

    # Resolve target path
    $target = $null
    if ($Path) {
        if (-not (Test-Path $Path)) { Write-Host "[ERROR] File not found: $Path" -ForegroundColor Red; return }
        $target = $Path
    } elseif ($VBE.Config) {
        $target = $VBE.Config.master_workbook
    }

    if (-not $target) {
        Write-Host "[ERROR] No workbook specified. Use: vbe open -Path `"path\to\workbook.xlsm`"" -ForegroundColor Red
        Write-Host "  Or create a vbe-auto-config.json in the toolkit directory." -ForegroundColor Gray
        return
    }

    vbe-log "Opening: $target" "SESSION"

    # Kill any existing Excel
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep 2

    # Start Excel with full automation control
    $VBE.Excel = New-Object -ComObject Excel.Application
    $VBE.Excel.Visible = $Visible.IsPresent
    $VBE.Excel.DisplayAlerts = $false
    $VBE.Excel.AutomationSecurity = 1
    $VBE.Excel.ScreenUpdating = $false
    $VBE.Excel.EnableEvents = (-not $NoEvents.IsPresent)
    $VBE.Excel.Interactive = $false  # Auto-dismiss ALL dialogs

    try {
        $VBE.Workbook = $VBE.Excel.Workbooks.Open($target, 0, $false)
    } catch {
        Write-Host "[ERROR] Failed to open: $($_.Exception.Message)" -ForegroundColor Red
        $VBE.Excel.Quit()
        $VBE.Excel = $null
        return
    }

    $VBE.SessionStart = Get-Date
    $VBE.CommandCount++

    $sheetCount = $VBE.Workbook.Sheets.Count
    $modCount = $VBE.Workbook.VBProject.VBComponents.Count
    $msg = "Opened: $($VBE.Workbook.Name) | $sheetCount sheets | $modCount modules"
    vbe-log $msg
    Write-Host $msg -ForegroundColor Green

    vbe-state-save
    return $VBE.Workbook
}

function vbe-save {
    if (-not $VBE.Workbook) { throw "No workbook open. Run 'vbe open' first." }
    vbe-log "Saving workbook" "SAVE"
    $VBE.Workbook.Save()
    $size = (Get-Item $VBE.Workbook.FullName).Length
    $msg = "Saved: $([math]::Round($size/1KB,1)) KB | Modified: $(Get-Date -Format 'HH:mm:ss')"
    vbe-log $msg
    Write-Host $msg -ForegroundColor Green
    vbe-state-save
}

function vbe-close {
    if ($VBE.Workbook) {
        vbe-log "Closing workbook" "SESSION"
        $VBE.Workbook.Close($false)
    }
    if ($VBE.Excel) {
        $VBE.Excel.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseCOMObject($VBE.Excel)
    }
    $VBE.Excel = $null
    $VBE.Workbook = $null
    vbe-log "Excel session closed" "SESSION"
    Write-Host "[VBE] Session closed" -ForegroundColor Yellow
    vbe-state-save
}

# ============================================================================
# MACRO EXECUTION
# ============================================================================

function vbe-macro {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [switch]$Silent,
        [switch]$Verbose
    )

    if (-not $VBE.Excel) { throw "Excel not open. Run 'vbe open' first." }

    vbe-log "Running macro: $Name" "MACRO"
    $VBE.CommandCount++

    try {
        $VBE.Excel.Interactive = $false
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $result = $VBE.Excel.Run($Name)
        $sw.Stop()

        if (-not $Silent) {
            Write-Host "[MACRO] $Name : SUCCESS ($($sw.ElapsedMilliseconds)ms)" -ForegroundColor Green
        }
        vbe-log "Macro ${Name}: SUCCESS in $($sw.ElapsedMilliseconds)ms"
        return @{ Success=$true; Duration=$sw.ElapsedMilliseconds; Result=$result }
    } catch {
        $err = $_.Exception.Message
        $VBE.LastError = $err
        if (-not $Silent) {
            Write-Host "[MACRO] $Name : FAILED - $err" -ForegroundColor Red
        }
        vbe-log "Macro ${Name}: FAILED - $err" "ERROR"
        $err | Out-File "$env:USERPROFILE\Desktop\vbe_macro_error.txt" -Encoding UTF8
        return @{ Success=$false; Error=$err }
    }
}

function vbe-macro-all {
    param([string]$ConfigMacroList)

    # Get macro list from config or default
    $macros = @()
    if ($VBE.Config -and $VBE.Config.macros -and $VBE.Config.macros.test_all) {
        $macros = @($VBE.Config.macros.test_all)
    }

    if ($macros.Count -eq 0) {
        # Auto-discover public macros from MAIN_MACROS and other modules
        if ($VBE.Workbook) {
            foreach ($comp in $VBE.Workbook.VBProject.VBComponents) {
                $cm = $comp.CodeModule
                for ($line = 1; $line -le $cm.CountOfLines; $line++) {
                    $text = $cm.Lines($line, 1).Trim()
                    if ($text -match '^Public\s+Sub\s+(\w+)') {
                        $macros += $matches[1]
                    }
                }
            }
            # Remove duplicates
            $macros = $macros | Select-Object -Unique
        }
    }

    if ($macros.Count -eq 0) {
        Write-Host "[VBE] No macros discovered. Add macros to config or ensure Public Sub declarations exist." -ForegroundColor Yellow
        return
    }

    Write-Host "`n=== Running $($macros.Count) Macros ===" -ForegroundColor Cyan
    $results = @()
    foreach ($macro in $macros) {
        $result = vbe-macro $macro -Silent
        $results += [PSCustomObject]@{
            Macro = $macro
            Status = if ($result.Success) { "PASS" } else { "FAIL" }
            Duration = if ($result.Duration) { "$($result.Duration)ms" } else { "N/A" }
            Error = if (-not $result.Success) { $result.Error } else { "-" }
        }
        Start-Sleep 0.3
    }

    $results | Format-Table -AutoSize
    $results | Export-Csv "$env:USERPROFILE\Desktop\vbe_macro_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation

    $passed = ($results | Where-Object { $_.Status -eq "PASS" }).Count
    $total = $results.Count
    vbe-log "Macro test: $passed/$total passed" "TEST"
    Write-Host "`nSummary: $passed/$total passed" -ForegroundColor $(if ($passed -eq $total) { "Green" } else { "Yellow" })
}

# ============================================================================
# COMPILATION
# ============================================================================

function vbe-compile {
    if (-not $VBE.Excel) { throw "Excel not open." }

    vbe-log "Compiling VBA project" "COMPILE"
    $VBE.CommandCount++

    try {
        $VBE.Excel.VBE.CommandBars("Menu Bar").Controls("Debug").Controls("Compile VBAProject").Execute()
        vbe-log "Compile: OK"
        Write-Host "[VBE] Compile: OK" -ForegroundColor Green
        return $true
    } catch {
        $err = $_.Exception.Message
        vbe-log "Compile: FAILED - $err" "ERROR"
        Write-Host "[VBE] Compile: FAILED - $err" -ForegroundColor Red
        return $false
    }
}

# ============================================================================
# MODULE MANAGEMENT
# ============================================================================

function vbe-modules {
    if (-not $VBE.Workbook) { throw "No workbook open." }

    $VBE.CommandCount++
    $modules = @()
    foreach ($comp in $VBE.Workbook.VBProject.VBComponents) {
        $type = switch ($comp.Type) { 1 { "Module" } 2 { "Class" } 3 { "Form" } 100 { "Document" } default { "Unknown" } }
        $modules += [PSCustomObject]@{
            Name = $comp.Name
            Type = $type
            Lines = $comp.CodeModule.CountOfLines
        }
    }

    $modules | Sort-Object Type, Name | Format-Table -AutoSize
    vbe-log "Listed $($modules.Count) modules"

    $totalLines = ($modules | Measure-Object -Property Lines -Sum).Sum
    Write-Host "Total: $($modules.Count) modules, $totalLines lines" -ForegroundColor Cyan
}

function vbe-module-read {
    param([Parameter(Mandatory=$true)][string]$ModuleName, [string]$OutputPath)

    if (-not $VBE.Workbook) { throw "No workbook open." }

    try {
        $comp = $VBE.Workbook.VBProject.VBComponents.Item($ModuleName)
        $code = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)

        if ($OutputPath) {
            $code | Out-File $OutputPath -Encoding UTF8
            Write-Host "[VBE] Exported $ModuleName to $OutputPath" -ForegroundColor Green
        } else {
            Write-Host "[VBE] === $ModuleName ($($comp.CodeModule.CountOfLines) lines) ===" -ForegroundColor Cyan
            Write-Host $code
            Write-Host "[VBE] === END ===" -ForegroundColor Cyan
        }
        vbe-log "Read module: $ModuleName ($($comp.CodeModule.CountOfLines) lines)"
        return $code
    } catch {
        Write-Host "[VBE] Module not found: $ModuleName" -ForegroundColor Red
        return $null
    }
}

function vbe-export-all {
    param([string]$DestPath)

    if (-not $VBE.Workbook) { throw "No workbook open." }

    $dest = if ($DestPath) { $DestPath } else { "$VBEToolkitDir\export_$(Get-Date -Format 'yyyyMMdd_HHmmss')" }
    New-Item -ItemType Directory -Path $dest -Force | Out-Null

    foreach ($comp in $VBE.Workbook.VBProject.VBComponents) {
        $ext = switch ($comp.Type) { 1 { ".bas" } 2 { ".cls" } 3 { ".frm" } 100 { ".cls" } default { ".txt" } }
        $comp.Export("$dest\$($comp.Name)$ext")
    }

    $count = $VBE.Workbook.VBProject.VBComponents.Count
    vbe-log "Exported $count modules to $dest"
    Write-Host "[VBE] Exported $count modules to $dest" -ForegroundColor Green
    return $dest
}

# ============================================================================
# SHEET NAVIGATION
# ============================================================================

function vbe-sheets {
    if (-not $VBE.Workbook) { throw "No workbook open." }

    $sheets = @()
    foreach ($ws in $VBE.Workbook.Sheets) {
        $lastRow = $ws.Cells($ws.Rows.Count, "A").End(-4162).Row
        $lastCol = $ws.Cells(1, $ws.Columns.Count).End(-4159).Column
        $sheets += [PSCustomObject]@{
            Name = $ws.Name
            Rows = $lastRow
            Columns = $lastCol
            Visible = if ($ws.Visible -eq -1) { "Visible" } else { "Hidden" }
        }
    }

    $sheets | Format-Table -AutoSize
    vbe-log "Listed $($sheets.Count) sheets"
}

function vbe-sheet {
    param([Parameter(Mandatory=$true)][string]$Name, [int]$MaxRows = 5)

    if (-not $VBE.Workbook) { throw "No workbook open." }

    try {
        $ws = $VBE.Workbook.Sheets.Item($Name)
    } catch {
        Write-Host "[VBE] Sheet not found: $Name" -ForegroundColor Red
        return
    }

    $lastRow = $ws.Cells($ws.Rows.Count, "A").End(-4162).Row
    $lastCol = $ws.Cells(1, $ws.Columns.Count).End(-4159).Column

    Write-Host "[VBE] Sheet: $Name | Rows: $lastRow | Columns: $lastCol" -ForegroundColor Cyan

    # Show headers
    $headers = @()
    for ($c = 1; $c -le [Math]::Min($lastCol, 20); $c++) {
        $val = $ws.Cells(1, $c).Value
        if ($val) { $headers += $val }
    }
    Write-Host "Headers: $($headers -join ' | ')" -ForegroundColor Yellow

    # Show first N data rows
    $showRows = [Math]::Min($MaxRows, $lastRow - 1)
    if ($showRows -gt 0) {
        for ($r = 2; $r -le ($showRows + 1); $r++) {
            $row = @()
            for ($c = 1; $c -le [Math]::Min($lastCol, 10); $c++) {
                $row += $ws.Cells($r, $c).Value
            }
            Write-Host "Row $($r-1): $($row -join ' | ')" -ForegroundColor Gray
        }
    }

    vbe-log "Inspected sheet: $Name ($lastRow rows, $lastCol cols)"
}

# ============================================================================
# STATE & SNAPSHOT
# ============================================================================

function vbe-snapshot {
    $snapshot = [PSCustomObject]@{
        Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Project = if ($VBE.Config) { $VBE.Config.project_name } else { "Unknown" }
        Version = if ($VBE.Config) { $VBE.Config.version } else { "N/A" }
        Workbook = if ($VBE.Workbook) { $VBE.Workbook.Name } else { "None" }
        FilePath = if ($VBE.Workbook) { $VBE.Workbook.FullName } else { "N/A" }
        Sheets = if ($VBE.Workbook) { $VBE.Workbook.Sheets.Count } else { 0 }
        Modules = if ($VBE.Workbook) { $VBE.Workbook.VBProject.VBComponents.Count } else { 0 }
        FileSize = if ($VBE.Workbook) { "$([math]::Round((Get-Item $VBE.Workbook.FullName).Length/1KB,1)) KB" } else { "N/A" }
        Modified = if ($VBE.Workbook) { (Get-Item $VBE.Workbook.FullName).LastWriteTime.ToString('HH:mm:ss') } else { "N/A" }
        SessionStart = if ($VBE.SessionStart) { $VBE.SessionStart.ToString('HH:mm:ss') } else { "N/A" }
        Commands = $VBE.CommandCount
        SessionDuration = if ($VBE.SessionStart) { "$([math]::Round((Get-Date - $VBE.SessionStart).TotalMinutes, 1)) min" } else { "N/A" }
        LastError = if ($VBE.LastError) { $VBE.LastError } else { "None" }
    }

    $snapshot | Format-List
    vbe-log "Snapshot taken"
    return $snapshot
}

# ============================================================================
# DIRECT COM ACCESS (for custom operations)
# ============================================================================

function vbe-excel {
    if (-not $VBE.Excel) { throw "Excel not open. Run 'vbe open' first." }
    Write-Host "[VBE] Excel COM object available as `$VBE.Excel" -ForegroundColor Cyan
    Write-Host "[VBE] Workbook available as `$VBE.Workbook" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host "  `$VBE.Excel.Run(`"MacroName`")" -ForegroundColor Gray
    Write-Host "  `$VBE.Workbook.Sheets.Count" -ForegroundColor Gray
    Write-Host "  `$VBE.Workbook.VBProject.VBComponents" -ForegroundColor Gray
    return $VBE
}

# ============================================================================
# CLI INTERFACE
# ============================================================================

function vbe {
    param(
        [Parameter(Position=0)]
        [ValidateSet("open","close","save","compile","macro","macro-all","modules","module-read","export-all","sheets","sheet","snapshot","log","state","excel","help")]
        [string]$Command,
        [Parameter(Position=1,ValueFromRemainingArguments=$true)]
        [string[]]$Args
    )

    switch ($Command) {
        "open" {
            $visible = $Args -contains "-Visible"
            $noEvents = $Args -contains "-NoEvents"
            $path = ($Args | Where-Object { $_ -notmatch "^-" }) | Select-Object -First 1
            vbe-open -Visible:$visible -NoEvents:$noEvents -Path:$path
        }
        "close" { vbe-close }
        "save" { vbe-save }
        "compile" { vbe-compile }
        "macro" {
            if (-not $Args[0]) { Write-Host "Usage: vbe macro <MacroName>" -ForegroundColor Yellow; return }
            vbe-macro $Args[0] -Verbose:($Args -contains "-Verbose")
        }
        "macro-all" { vbe-macro-all }
        "modules" { vbe-modules }
        "module-read" {
            if (-not $Args[0]) { Write-Host "Usage: vbe module-read <ModuleName> [-OutputPath <path>]" -ForegroundColor Yellow; return }
            $out = $null
            $idx = $Args.IndexOf("-OutputPath")
            if ($idx -ge 0 -and $idx -lt ($Args.Count - 1)) { $out = $Args[$idx + 1] }
            vbe-module-read -ModuleName $Args[0] -OutputPath $out
        }
        "export-all" {
            $path = ($Args | Where-Object { $_ -notmatch "^-" }) | Select-Object -First 1
            vbe-export-all -DestPath $path
        }
        "sheets" { vbe-sheets }
        "sheet" {
            if (-not $Args[0]) { Write-Host "Usage: vbe sheet <SheetName>" -ForegroundColor Yellow; return }
            $maxRows = 5
            $idx = $Args.IndexOf("-MaxRows")
            if ($idx -ge 0 -and $idx -lt ($Args.Count - 1)) { $maxRows = [int]$Args[$idx + 1] }
            vbe-sheet $Args[0] -MaxRows:$maxRows
        }
        "snapshot" { vbe-snapshot }
        "log" {
            $lines = if ($Args[0] -match "^\d+$") { [int]$Args[0] } else { 20 }
            if (Test-Path $VBELogPath) { Get-Content $VBELogPath -Tail $lines }
        }
        "state" { vbe-state-load }
        "excel" { vbe-excel }
        "help" {
            Write-Host @"

=== VBE Auto — Universal Excel VBE Control Suite ===

Usage: vbe <command> [arguments]

Session:
  vbe open [-Visible] [-Path "file.xlsm"]  Open workbook (auto-discovers from config)
  vbe open -ListProjects                    List registered projects
  vbe close                                 Close Excel session
  vbe save                                  Save current workbook

Compilation:
  vbe compile                               Compile VBA project

Macros:
  vbe macro <Name>                          Run a specific macro
  vbe macro-all                             Test all discovered macros

Modules:
  vbe modules                               List all VBA modules
  vbe module-read <Name> [-OutputPath p]    Read module code
  vbe export-all [<path>]                   Export all modules to folder

Sheets:
  vbe sheets                                List all sheets with dimensions
  vbe sheet <Name> [-MaxRows N]             Inspect sheet content

State:
  vbe snapshot                              Show current session state
  vbe log [<lines>]                         Show recent operations log
  vbe state                                 Load last saved state
  vbe excel                                 Access raw COM objects

Examples:
  vbe open
  vbe open -Path "C:\project\MyApp.xlsm"
  vbe macro GenerateDemoData
  vbe compile
  vbe snapshot
  vbe macro-all
  vbe sheet MOUVEMENTS -MaxRows 10
  vbe close

"@ -ForegroundColor Cyan
        }
        default {
            Write-Host "Unknown command: $Command. Run 'vbe help' for usage." -ForegroundColor Red
        }
    }
}

# ============================================================================
# AUTO-LOAD MESSAGE
# ============================================================================

$projName = if ($VBEConfig) { $VBEConfig.project_name } else { "No project config found" }
Write-Host "`n=== VBE Auto Control Suite Loaded ===" -ForegroundColor Green
Write-Host "Project: $projName" -ForegroundColor Cyan
Write-Host "Type 'vbe help' for commands" -ForegroundColor Cyan
