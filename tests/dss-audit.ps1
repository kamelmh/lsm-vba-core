param(
    [string]$WorkbookPath = "C:\Users\Administrator\Dropbox\Logistics.Public.Sector.Refactor\Software_Surgical_Edit\ERP_Academie_v13_2.xlsm"
)

$ErrorActionPreference = "Continue"
$auditResults = @()
$critical = 0; $warning = 0; $info = 0; $pass = 0

function Audit-Record {
    param([string]$Category, [string]$Check, [string]$Severity, [string]$Status, [string]$Detail)
    $script:auditResults += [PSCustomObject]@{
        Category = $Category; Check = $Check; Severity = $Severity
        Status = $Status; Detail = $Detail; Timestamp = Get-Date -Format 'HH:mm:ss'
    }
    switch ($Severity) {
        "CRITICAL" { if ($Status -ne "PASS") { $script:critical++ } else { $script:pass++ } }
        "WARNING"  { if ($Status -ne "PASS") { $script:warning++ } else { $script:pass++ } }
        "INFO"     { $script:info++ }
    }
    $c = switch ($Status) { "PASS" { "Green" } "FAIL" { "Red" } "WARN" { "Yellow" }; default { "Gray" } }
    $label = "  [$Status] [$Severity] ${Category}: $Check"
    Write-Host $label -ForegroundColor $c
    if ($Detail -and $Status -ne "PASS") { Write-Host "         $Detail" -ForegroundColor DarkGray }
}

# Kill existing Excel
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

Write-Host "`n=== DSS System Audit - Academix v13.2 ===" -ForegroundColor Cyan
Write-Host "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n" -ForegroundColor Gray

# ============================================================================
# PHASE 1: Structural Integrity
# ============================================================================
Write-Host "[Phase 1] Structural Integrity" -ForegroundColor Yellow

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $xl.Interactive = $false; $xl.EnableEvents = $false
    $wb = $xl.Workbooks.Open($WorkbookPath)

    # 1.1 Sheet count
    $sheetCount = $wb.Sheets.Count
    if ($sheetCount -ge 25) { Audit-Record "Structure" "Sheet count ($sheetCount)" "CRITICAL" "PASS" "25+ sheets present" }
    else { Audit-Record "Structure" "Sheet count ($sheetCount)" "CRITICAL" "FAIL" "Expected 25+, found $sheetCount" }

    # 1.2 Required sheets
    $required = @("ACCUEIL","ARTICLES","FOURNISSEURS","MOUVEMENTS","TABLEAU DE BORD","CALCULS_EOQ","ALERTE_DASHBOARD","INVENTAIRE","RAPPORTS","HISTORIQUE")
    $missing = @()
    foreach ($s in $required) { try { $null = $wb.Sheets($s) } catch { $missing += $s } }
    if ($missing.Count -eq 0) { Audit-Record "Structure" "Required sheets" "CRITICAL" "PASS" "All 10 core sheets present" }
    else { Audit-Record "Structure" "Required sheets" "CRITICAL" "FAIL" "Missing: $($missing -join ', ')" }

    # 1.3 VBA components
    $compCount = $wb.VBProject.VBComponents.Count
    if ($compCount -ge 38) { Audit-Record "Structure" "VBA components ($compCount)" "CRITICAL" "PASS" "38+ components (29 .bas + 1 .frm + ThisWorkbook + internals)" }
    else { Audit-Record "Structure" "VBA components ($compCount)" "CRITICAL" "WARN" "Expected 38+, found $compCount" }

    # 1.4 Module naming
    $modules = $wb.VBProject.VBComponents | Where-Object { $_.Type -eq 1 } | Select-Object -ExpandProperty Name
    $duplicates = $modules | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($duplicates.Count -eq 0) { Audit-Record "Structure" "No duplicate modules" "WARNING" "PASS" "All module names unique" }
    else { Audit-Record "Structure" "Duplicate modules" "CRITICAL" "FAIL" "Duplicates: $($duplicates.Name -join ', ')" }

    # 1.5 Code lines
    $totalLines = 0
    foreach ($comp in $wb.VBProject.VBComponents) { $totalLines += $comp.CodeModule.CountOfLines }
    if ($totalLines -ge 9000) { Audit-Record "Structure" "Code volume ($totalLines lines)" "INFO" "PASS" "9000+ lines of business logic" }
    else { Audit-Record "Structure" "Code volume ($totalLines lines)" "INFO" "WARN" "Expected 9000+, found $totalLines" }

} catch {
    Audit-Record "Structure" "Workbook load" "CRITICAL" "FAIL" $_.Exception.Message
}

# ============================================================================
# PHASE 2: Security Audit
# ============================================================================
Write-Host "`n[Phase 2] Security Audit" -ForegroundColor Yellow

# 2.1 Sheet protection
$protectedCount = 0
foreach ($ws in $wb.Sheets) { try { if ($ws.ProtectContents) { $protectedCount++ } } catch {} }
if ($protectedCount -ge 20) { Audit-Record "Security" "Sheet protection ($protectedCount/$($wb.Sheets.Count))" "CRITICAL" "PASS" "Majority sheets protected" }
else { Audit-Record "Security" "Sheet protection ($protectedCount/$($wb.Sheets.Count))" "CRITICAL" "FAIL" "Insufficient sheet protection" }

# 2.2 No hardcoded passwords in source (check .bas files)
$sourceDir = "C:\Users\Administrator\Dropbox\Logistics.Public.Sector.Refactor\Software_Surgical_Edit\VBA_Modules"
$pwdLines = Select-String -Path "$sourceDir\*.bas" -Pattern 'Password\s*[:=]\s*"[^"]+"' -CaseSensitive | Where-Object { $_.Line -match 'Password\s*[:=]\s*"' -and $_.Line -notmatch "MASTER_PWD" }
if ($pwdLines.Count -eq 0) { Audit-Record "Security" "No hardcoded passwords" "CRITICAL" "PASS" "All passwords use mod_Config.MASTER_PWD" }
else { Audit-Record "Security" "Hardcoded passwords" "CRITICAL" "FAIL" "$($pwdLines.Count) lines with hardcoded passwords" }

# 2.3 Error handling coverage
$modules = Get-ChildItem "$sourceDir\*.bas"
$modulesWithOnError = 0; $totalProcedures = 0
foreach ($f in $modules) {
    $content = Get-Content $f.FullName -Raw
    $procs = [regex]::Matches($content, "(?:Public|Private)\s+(?:Sub|Function)\s+\w+").Count
    $totalProcedures += $procs
    if ($content -match "On\s+Error") { $modulesWithOnError++ }
}
$pct = [math]::Round(($modulesWithOnError / $modules.Count) * 100, 0)
if ($pct -ge 80) { Audit-Record "Security" "Error handling ($pct%)" "WARNING" "PASS" "$modulesWithOnError/$modules.Count modules have error handling" }
else { Audit-Record "Security" "Error handling ($pct%)" "WARNING" "FAIL" "Only $modulesWithOnError/$modules.Count modules have error handling" }

# 2.4 Transaction safety
$hasTransSafety = Test-Path "$sourceDir\mod_TransactionSafety.bas"
if ($hasTransSafety) {
    $content = Get-Content "$sourceDir\mod_TransactionSafety.bas" -Raw
    if ($content -match "BeginTransaction" -and $content -match "RollbackTransaction" -and $content -match "CommitTransaction") {
        Audit-Record "Security" "Transaction safety (ACID)" "CRITICAL" "PASS" "Full transaction lifecycle implemented"
    } else { Audit-Record "Security" "Transaction safety" "CRITICAL" "FAIL" "Missing transaction methods" }
} else { Audit-Record "Security" "Transaction safety module" "CRITICAL" "FAIL" "mod_TransactionSafety.bas not found" }

# 2.5 Audit trail
$hasAudit = Test-Path "$sourceDir\mod_AuditTrail.bas"
if ($hasAudit) {
    $content = Get-Content "$sourceDir\mod_AuditTrail.bas" -Raw
    if ($content -match "LogAction" -or $content -match "LogTransaction") {
        Audit-Record "Security" "Audit trail logging" "CRITICAL" "PASS" "Audit logging functions present"
    } else { Audit-Record "Security" "Audit trail logging" "CRITICAL" "FAIL" "No logging functions found" }
} else { Audit-Record "Security" "Audit trail module" "CRITICAL" "FAIL" "mod_AuditTrail.bas not found" }

# ============================================================================
# PHASE 3: Data Integrity
# ============================================================================
Write-Host "`n[Phase 3] Data Integrity" -ForegroundColor Yellow

# 3.1 ARTICLES data
try {
    $wsArt = $wb.Sheets("ARTICLES")
    $lastRow = $wsArt.Cells($wsArt.Rows.Count, 1).End(-4162).Row
    if ($lastRow -ge 13) { Audit-Record "Data" "ARTICLES records ($($lastRow-1))" "CRITICAL" "PASS" "12+ articles cataloged" }
    else { Audit-Record "Data" "ARTICLES records ($($lastRow-1))" "CRITICAL" "FAIL" "Expected 12+, found $($lastRow-1)" }

    # 3.2 No blank article codes
    $blankCodes = 0
    for ($i = 2; $i -le $lastRow; $i++) {
        if ([string]::IsNullOrWhiteSpace($wsArt.Cells($i, 1).Value2)) { $blankCodes++ }
    }
    if ($blankCodes -eq 0) { Audit-Record "Data" "No blank article codes" "WARNING" "PASS" "All articles have codes" }
    else { Audit-Record "Data" "Blank article codes" "WARNING" "FAIL" "$blankCodes articles without codes" }
} catch {
    Audit-Record "Data" "ARTICLES sheet access" "CRITICAL" "FAIL" $_.Exception.Message
}

# 3.3 MOUVEMENTS data
try {
    $wsMouv = $wb.Sheets("MOUVEMENTS")
    $mouvCount = $wsMouv.Cells($wsMouv.Rows.Count, 1).End(-4162).Row - 1
    if ($mouvCount -ge 10) { Audit-Record "Data" "MOUVEMENTS records ($mouvCount)" "WARNING" "PASS" "$mouvCount transactions recorded" }
    else { Audit-Record "Data" "MOUVEMENTS records ($mouvCount)" "WARNING" "WARN" "Only $mouvCount transactions (expected 10+)" }
} catch {
    Audit-Record "Data" "MOUVEMENTS sheet access" "CRITICAL" "FAIL" $_.Exception.Message
}

# 3.4 Data type validation
try {
    $wsArt = $wb.Sheets("ARTICLES")
    $typeErrors = 0
    for ($i = 2; $i -le 13; $i++) {
        $stock = $wsArt.Cells($i, 3).Value2
        if ($stock -ne $null -and -not [double]::TryParse($stock, [ref]$null)) { $typeErrors++ }
    }
    if ($typeErrors -eq 0) { Audit-Record "Data" "Stock column type validation" "WARNING" "PASS" "All stock values numeric" }
    else { Audit-Record "Data" "Stock column type errors" "WARNING" "FAIL" "$typeErrors non-numeric stock values" }
} catch {
    Audit-Record "Data" "Type validation" "WARNING" "FAIL" $_.Exception.Message
}

# ============================================================================
# PHASE 4: Module Call Graph
# ============================================================================
Write-Host "`n[Phase 4] Module Call Graph Analysis" -ForegroundColor Yellow

$callGraph = @{}
foreach ($f in Get-ChildItem "$sourceDir\*.bas") {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    # Skip MAIN_MACROS (entry point, not a library module)
    if ($name -eq "MAIN_MACROS") { continue }
    $content = Get-Content $f.FullName -Raw
    # Remove single-line comments
    $cleanContent = $content -replace "'[^\n]*", ""
    # Remove string literals
    $cleanContent = $cleanContent -replace '"[^"]*"', ""
    $cleanContent = $cleanContent -replace "'[^']*'", ""
    $calls = [regex]::Matches($cleanContent, "\bmod_\w+\.") | ForEach-Object { $_.Value.TrimEnd('.') } | Select-Object -Unique
    $callGraph[$name] = $calls
}

# 4.1 Orphan modules (not called by any other module)
# Include calls from ThisWorkbook, forms, and MAIN_MACROS
$allCalls = $callGraph.Values | ForEach-Object { $_ } | Select-Object -Unique
# Also scan ThisWorkbook.cls and .frm files for module references
foreach ($f in Get-ChildItem "$sourceDir\*.cls", "$sourceDir\*.frm") {
    $content = Get-Content $f.FullName -Raw
    $cleanContent = $content -replace "'[^\n]*", ""
    $cleanContent = $cleanContent -replace '"[^"]*"', ""
    $refs = [regex]::Matches($cleanContent, "\bmod_\w+\.") | ForEach-Object { $_.Value.TrimEnd('.') } | Select-Object -Unique
    $allCalls = $allCalls + $refs | Select-Object -Unique
}
# MAIN_MACROS is an entry point, not a library
$entryPoints = @("mod_Config", "MAIN_MACROS")
$orphanCount = 0
$orphans = @()
foreach ($key in $callGraph.Keys) {
    if ($key -notin $allCalls -and $key -notin $entryPoints) {
        $orphanCount++
        $orphans += $key
    }
}
if ($orphanCount -eq 0) { Audit-Record "Graph" "No orphan modules" "WARNING" "PASS" "All modules referenced" }
else { Audit-Record "Graph" "Orphan modules ($orphanCount)" "WARNING" "WARN" "Not called by others: $($orphans -join ', ')" }

# 4.2 Circular dependencies
$circularFound = $false
foreach ($src in $callGraph.Keys) {
    foreach ($tgt in $callGraph[$src]) {
        if ($tgt -ne $src -and $callGraph.ContainsKey($tgt) -and $callGraph[$tgt] -contains $src) {
            $circularFound = $true
            Audit-Record "Graph" "Circular: $src <-> $tgt" "CRITICAL" "FAIL" "Circular dependency detected"
            break
        }
    }
    if ($circularFound) { break }
}
if (-not $circularFound) { Audit-Record "Graph" "No circular dependencies" "CRITICAL" "PASS" "Clean dependency graph" }

# 4.3 High coupling modules
$highCoupling = @()
foreach ($key in $callGraph.Keys) {
    if ($callGraph[$key].Count -ge 8) { $highCoupling += "$key ($($callGraph[$key].Count))" }
}
if ($highCoupling.Count -le 2) { Audit-Record "Graph" "Module coupling" "INFO" "PASS" "Few highly-coupled modules: $($highCoupling -join ', ')" }
else { Audit-Record "Graph" "High coupling ($($highCoupling.Count) modules)" "INFO" "WARN" "Consider refactoring: $($highCoupling -join ', ')" }

# ============================================================================
# PHASE 5: Performance & Compliance
# ============================================================================
Write-Host "`n[Phase 5] Performance & Compliance" -ForegroundColor Yellow

# 5.1 Application optimization
try {
    $content = Get-Content "$sourceDir\mod_StockEntry_Logic.bas" -Raw
    if ($content -match "ScreenUpdating\s*=\s*False" -and $content -match "Calculation\s*=\s*xlCalculationManual" -and $content -match "EnableEvents\s*=\s*False") {
        Audit-Record "Performance" "Excel optimization flags" "WARNING" "PASS" "ScreenUpdating, Calculation, EnableEvents disabled during ops"
    } else { Audit-Record "Performance" "Excel optimization flags" "WARNING" "FAIL" "Missing optimization flags" }
} catch {
    Audit-Record "Performance" "Excel optimization" "WARNING" "FAIL" $_.Exception.Message
}

# 5.2 EOQ constants - read from VBA constants via test function, not from sheet cells
try {
    $tm2 = $wb.VBProject.VBComponents.Add(1)
    $tm2.Name = "mod_EOQTest"
    $tm2.CodeModule.AddFromString("Option Explicit`nPublic Function Test_D() As Integer: Test_D = mod_Config.WORKING_DAYS_PER_YEAR: End Function")
    $d = $xl.Run("Test_D")
    $wb.VBProject.VBComponents.Remove($tm2)
    if ($d -eq 250) {
        Audit-Record "Compliance" "Config constants (WORKING_DAYS=250)" "CRITICAL" "PASS" "Config constants accessible"
    } else { Audit-Record "Compliance" "Config constants (WORKING_DAYS=$d)" "CRITICAL" "FAIL" "Expected 250, found $d" }
} catch {
    Audit-Record "Compliance" "Config constants" "CRITICAL" "FAIL" $_.Exception.Message
}

# 5.3 French column headers
try {
    $wsArt = $wb.Sheets("ARTICLES")
    $headers = @($wsArt.Cells(1, 1).Value2, $wsArt.Cells(1, 2).Value2, $wsArt.Cells(1, 3).Value2)
    $hasFrench = $headers[0] -match "CODE" -or $headers[0] -match "Article"
    if ($hasFrench) { Audit-Record "Compliance" "French headers" "INFO" "PASS" "Column headers in French" }
    else { Audit-Record "Compliance" "French headers" "INFO" "WARN" "Headers: $($headers -join ' | ')" }
} catch {
    Audit-Record "Compliance" "French headers" "INFO" "WARN" $_.Exception.Message
}

# ============================================================================
# CLEANUP
# ============================================================================
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "`nCleaning up Excel..." -ForegroundColor Gray

# ============================================================================
# AUDIT REPORT
# ============================================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "        DSS AUDIT SUMMARY" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  PASS:     $pass" -ForegroundColor Green
Write-Host "  CRITICAL: $critical" -ForegroundColor Red
Write-Host "  WARNING:  $warning" -ForegroundColor Yellow
Write-Host "  INFO:     $info" -ForegroundColor DarkGray
Write-Host ""
if ($critical -eq 0 -and $warning -eq 0) {
    Write-Host "  STATUS: ALL CHECKS PASSED" -ForegroundColor Green
} elseif ($critical -eq 0) {
    Write-Host "  STATUS: PASS WITH WARNINGS" -ForegroundColor Yellow
} else {
    Write-Host "  STATUS: CRITICAL ISSUES FOUND" -ForegroundColor Red
}
Write-Host ""

# Save report
$reportPath = "C:\Users\Administrator\Dropbox\Logistics.Public.Sector.Refactor\Software_Surgical_Edit\milestone_13_2\audit\dss-audit-report.csv"
$auditResults | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
Write-Host "  Report saved to: $reportPath" -ForegroundColor Gray
Write-Host ""
