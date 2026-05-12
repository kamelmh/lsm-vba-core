param(
    [string]$WorkbookPath = "C:\Users\Administrator\Dropbox\Logistics.Public.Sector.Refactor\Software_Surgical_Edit\ERP_Academie_v13_2.xlsm"
)

$ErrorActionPreference = "Stop"

Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.ScreenUpdating = $false
$excel.Interactive = $false

$allResults = @()

function Test-Macro {
    param([string]$Name, [string]$Macro, [int]$TimeoutMs = 15000)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $status = "FAIL"
    $error = ""
    try {
        $result = $null
        # COM objects can't cross PowerShell runspace boundaries — run inline
        $completed = $true
        try {
            $result = $excel.Run($Macro)
            $status = "PASS"
        } catch {
            $status = "FAIL"
            $error = $_.Exception.Message
        }
        if ($sw.ElapsedMilliseconds -gt $TimeoutMs) {
            $status = "TIMEOUT"
            $error = "Exceeded ${TimeoutMs}ms"
        }
    } catch {
        $status = "FAIL"
        $error = $_.Exception.Message
    }
    $sw.Stop()
    $script:allResults += [PSCustomObject]@{ Name = $Name; Status = $status; Error = $error; Ms = $sw.ElapsedMilliseconds }
    return $status
}

Write-Host "`n=== Macro Test Suite - ERP v13.2 ===" -ForegroundColor Cyan

try {
    $wb = $excel.Workbooks.Open($WorkbookPath)
    Write-Host "Workbook opened" -ForegroundColor Green
    Write-Host "  Sheets: $($wb.Sheets.Count)" -ForegroundColor Gray
    Write-Host "  VBA Components: $($wb.VBProject.VBComponents.Count)" -ForegroundColor Gray
} catch {
    Write-Host "Failed to open: $_" -ForegroundColor Red
    exit 1
}

$tm = $wb.VBProject.VBComponents.Add(1)
$tm.Name = "mod_Tests"

$code = @"
Option Explicit

Public Function Test_APP_VERSION() As String
    Test_APP_VERSION = mod_Config.APP_VERSION
End Function

Public Function Test_WORKING_DAYS() As Long
    Test_WORKING_DAYS = mod_Config.WORKING_DAYS_PER_YEAR
End Function

Public Function Test_GetSafetyStock() As Double
    Test_GetSafetyStock = mod_StockEngine.GetSafetyStock("ART-001")
End Function

Public Function Test_GetArticleStock() As Double
    Test_GetArticleStock = mod_StockEngine.GetArticleStock("ART-001")
End Function

Public Function Test_SyncMetrics() As Boolean
    Call mod_SyncBridge.SyncMetricsFromLedger
    Test_SyncMetrics = True
End Function

Public Function Test_GetSkuMetrics() As String
    Test_GetSkuMetrics = mod_SyncBridge.GetSkuMetrics("ART-001")
End Function

Public Function Test_GenerateVerifyCode() As String
    Test_GenerateVerifyCode = mod_Utilities.GenerateVerifyCode("TEST-DOC-001")
End Function

Public Function Test_SafeVal() As Double
    Test_SafeVal = mod_Utilities.SafeVal("1234.56")
End Function

Public Function Test_RestoreHeaders() As Boolean
    Call mod_Utilities.RestoreMouvementsHeaders(silent:=True)
    Test_RestoreHeaders = True
End Function

Public Function Test_GetAllArticles() As Variant
    Test_GetAllArticles = mod_Forecasting.GetAllArticles()
End Function
"@

$tm.CodeModule.AddFromString($code)
Write-Host "Test module injected`n" -ForegroundColor Gray

$tests = @(
    @{Name="APP_VERSION"; Macro="Test_APP_VERSION"; Timeout=10000},
    @{Name="WORKING_DAYS"; Macro="Test_WORKING_DAYS"; Timeout=10000},
    @{Name="GetSafetyStock"; Macro="Test_GetSafetyStock"; Timeout=10000},
    @{Name="GetArticleStock"; Macro="Test_GetArticleStock"; Timeout=10000},
    @{Name="SyncMetrics"; Macro="Test_SyncMetrics"; Timeout=15000},
    @{Name="GetSkuMetrics"; Macro="Test_GetSkuMetrics"; Timeout=10000},
    @{Name="GenerateVerifyCode"; Macro="Test_GenerateVerifyCode"; Timeout=10000},
    @{Name="SafeVal"; Macro="Test_SafeVal"; Timeout=10000},
    @{Name="RestoreHeaders"; Macro="Test_RestoreHeaders"; Timeout=10000},
    @{Name="GetAllArticles"; Macro="Test_GetAllArticles"; Timeout=15000}
)

foreach ($t in $tests) {
    $result = Test-Macro -Name $t.Name -Macro $t.Macro -TimeoutMs $t.Timeout
    $c = if ($result -eq "PASS") { "Green" } elseif ($result -eq "TIMEOUT") { "Yellow" } else { "Red" }
    Write-Host "  [$result] $($t.Name)" -ForegroundColor $c
}

Write-Host "`nCleaning up..." -ForegroundColor Gray
$wb.VBProject.VBComponents.Remove($tm)
$wb.Save()
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$pass = ($allResults | Where-Object { $_.Status -eq "PASS" }).Count
$fail = ($allResults | Where-Object { $_.Status -ne "PASS" }).Count

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  Passed: $pass  |  Failed: $fail  |  Total: $($allResults.Count)" -ForegroundColor White
Write-Host "============================================" -ForegroundColor Cyan
