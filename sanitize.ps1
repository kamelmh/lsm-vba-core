# Sanitize VBA source files for public release
# Removes sensitive data, replaces with generic placeholders

$ErrorActionPreference = "Stop"
$sourceDir = "C:\Users\Administrator\Dropbox\Logistics.Public.Sector.Refactor\Software_Surgical_Edit\VBA_Modules"
$publicDir = "C:\Users\Administrator\Dropbox\Logistics.Public.Sector.Refactor\Software_Surgical_Edit\milestone_13_2\public-lsm\src"

Write-Host "Sanitizing VBA modules for public release..." -ForegroundColor Cyan

# Files to include in public release
$includeFiles = @(
    "mod_Config.bas",
    "mod_StockEngine.bas",
    "mod_StockEntry_Logic.bas",
    "mod_Database.bas",
    "mod_TransactionSafety.bas",
    "mod_SyncBridge.bas",
    "mod_Forecasting.bas",
    "mod_Dashboard.bas",
    "mod_Reports.bas",
    "mod_ExportEngine.bas",
    "mod_ThemingEngine.bas",
    "mod_UI_Setup.bas",
    "mod_Navigation.bas",
    "mod_Localization.bas",
    "mod_Utilities.bas",
    "mod_SheetSetup.bas",
    "mod_SharedEnvironment.bas",
    "mod_AuditTrail.bas",
    "mod_ApprovalWorkflow.bas",
    "mod_QRCode.bas",
    "mod_ReceiptTag.bas",
    "mod_Analysis.bas",
    "mod_DemoData.bas",
    "mod_CSVImportExport.bas",
    "mod_Barcode.bas",
    "mod_InventoryReconciliation.bas",
    "mod_StockOutPredictor.bas",
    "mod_SupplierScorecard.bas",
    "mod_StockAging.bas",
    "mod_DataValidator.bas",
    "frmStockEntry.frm"
)

# Sanitization rules
$sanitizations = @(
    @{ Pattern = 'erp_secure_pwd_2026'; Replacement = '[YOUR_MASTER_PASSWORD]' },
    @{ Pattern = 'Mahi Kamel Abdelghani'; Replacement = 'LSM VBA Core' },
    @{ Pattern = 'CNEPD'; Replacement = 'Public Sector' },
    @{ Pattern = 'El Bayadh'; Replacement = '[CITY]' },
    @{ Pattern = 'Direction de l''Éducation'; Replacement = '[ORGANIZATION]' },
    @{ Pattern = 'Direction de l''Education'; Replacement = '[ORGANIZATION]' },
    @{ Pattern = 'Toner imprimante'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Rame papier'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Boîte archives'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Agrafeuse'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Stylos'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Registre'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Encre tampon'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Sous-Chemise'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Sous-chemise carton'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Chemise cartonn'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Rouleau papier fax'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Marqueur permanent'; Replacement = '[ARTICLE_DESC]' },
    @{ Pattern = 'Fournitures d''impression'; Replacement = '[CATEGORY]' },
    @{ Pattern = 'ENAP Alger'; Replacement = '[SUPPLIER_1]' },
    @{ Pattern = 'Bureautique Oran'; Replacement = '[SUPPLIER_2]' },
    @{ Pattern = 'Bureau Plus'; Replacement = '[SUPPLIER_3]' },
    @{ Pattern = 'Ministère de l''Éducation Nationale'; Replacement = '[MINISTRY]' },
    @{ Pattern = 'République Algérienne'; Replacement = '[COUNTRY]' },
    @{ Pattern = 'Académie'; Replacement = 'LSM' },
    @{ Pattern = 'Academie'; Replacement = 'LSM' },
    @{ Pattern = 'Academix'; Replacement = 'LSM' },
    @{ Pattern = 'ERP_Academie'; Replacement = 'LSM_Core' },
    @{ Pattern = 'v13.2'; Replacement = 'v1.0.0' },
    @{ Pattern = 'v13_2'; Replacement = 'v1_0_0' },
    @{ Pattern = '2026-05-07'; Replacement = '2026-01-01' },
    @{ Pattern = '2026-05-06'; Replacement = '2026-01-01' },
    @{ Pattern = '2026-05-04'; Replacement = '2026-01-01' }
)

foreach ($f in $includeFiles) {
    $srcPath = Join-Path $sourceDir $f
    $dstPath = Join-Path $publicDir $f
    
    if (-not (Test-Path $srcPath)) {
        Write-Host "  SKIP: $f (not found)" -ForegroundColor Yellow
        continue
    }
    
    $content = Get-Content $srcPath -Raw -Encoding UTF8
    
    foreach ($rule in $sanitizations) {
        $content = $content -replace [regex]::Escape($rule.Pattern), $rule.Replacement
    }
    
    # Set default password placeholder
    $content = $content -replace 'MASTER_PWD\s*=\s*".*?"', 'MASTER_PWD = "[YOUR_MASTER_PASSWORD]"'
    $content = $content -replace 'MASTER_PWD\s*=\s*''[^'']*''', 'MASTER_PWD = ''[YOUR_MASTER_PASSWORD]'''
    
    # Fix split-string institution name using Chr(201) for É
    $content = $content -replace '"Direction de l.*?" & Chr\(201\) & "ducation[^"]*"', '"[ORGANIZATION]"'
    
    $content | Out-File $dstPath -Encoding UTF8 -NoNewline
    $lineCount = ($content -split "`n").Count
    Write-Host "  SANITIZED: $f ($lineCount lines)" -ForegroundColor Green
}

# Copy test and audit scripts
Copy-Item -Force "C:\Users\Administrator\Dropbox\Logistics.Public.Sector.Refactor\Software_Surgical_Edit\milestone_13_2\tests\dss-audit.ps1" "$publicDir\..\tests\dss-audit.ps1"
Copy-Item -Force "C:\Users\Administrator\Dropbox\Logistics.Public.Sector.Refactor\Software_Surgical_Edit\test-macros.ps1" "$publicDir\..\tests\test-macros.ps1"
Write-Host "  COPIED: test-macros.ps1, dss-audit.ps1" -ForegroundColor Green

Write-Host "`nDone. Public modules saved to: $publicDir" -ForegroundColor Cyan
