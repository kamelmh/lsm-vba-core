# Changelog

All notable changes to LSM VBA Core will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [Unreleased]

## [1.4.0] — 2026-05-08

### Added
- `mod_UI_Setup.bas` — Programmatic dashboard button creation (16 buttons: stock entry, CSV, barcode, analysis modules, reports, KPIs)
- Inline field validation — red border + light red background on invalid inputs (HighlightFieldError/ClearFieldError)
- Type-to-search on `cmbArticle` dropdown (`fmStyleDropDownCombo`)
- Focus effects — blue border + light blue background on Enter, reset on Exit

### Changed
- `mod_ThemingEngine.bas` — Added `ApplyInputFocus`, `ApplyInputBlur`, `HighlightFieldError`, `ClearFieldError`, `ClearAllFieldErrors`
- `mod_StockEntry_Logic.bas` — Validation guards in `AddLineToGrid` and `CommitTransaction` call `HighlightFieldError` before `SetFocus`
- `frmStockEntry.frm` — Hover effects on 6 form buttons (blue/green/red shifts), 8 Enter/Exit focus handlers
- `README.md` — Architecture tree updated: `mod_UIEnhancements` → `mod_UI_Setup`
- `lsm-public-spec.xml` — 5 analysis modules added, `mod_UIEnhancements` moved to Excluded (dead code), `mod_UI_Setup` added to UI Framework
- `sanitize.ps1` — Include list updated: added `mod_UI_Setup.bas`, removed `mod_UIEnhancements.bas`

### Removed
- `mod_UIEnhancements.bas` — Dead code (386 lines, no callers) excluded from public release

## [1.3.0] — 2026-05-08

### Added
- `mod_InventoryReconciliation.bas` — Physical vs system inventory comparison with variance analysis
- `mod_StockOutPredictor.bas` — Average daily consumption + depletion date + CRITIQUE/ALERTE flags
- `mod_SupplierScorecard.bas` — Volume score + rating → A-D classification with color-coded sheet
- `mod_StockAging.bas` — Last movement → ACTIF/LENT/MORT categories with aging analysis
- `mod_DataValidator.bas` — Integrity scan of MOUVEMENTS/ARTICLES/FOURNISSEURS with error report

### Changed
- Public source now includes 31 modules (30 .bas + 1 .frm)
- CI runner fixed to `ubuntu-latest`

## [1.0.2] — 2026-05-08

### Added
- `mod_CSVImportExport.bas` — CSV import/export for MOUVEMENTS, ARTICLES, FOURNISSEURS with quoted fields and delimiter auto-detect
- `mod_Barcode.bas` — Keyboard-wedge barcode scanning with 12-article default mapping and STAGING_BUFFER-based custom map

### Changed
- CI workflow runner fixed to `ubuntu-latest`, `free-disk-space` step removed
- Public source now includes 26 modules (25 .bas + 1 .frm)
- All 26 sheets protected during build (previously had skip list)
- Repository URL references updated to `kamelmh/lsm-vba-core`

## [1.0.1] — 2026-05-08

### Added
- `mod_Analysis.bas` — ABC classification wrapper module
- `mod_DemoData.bas` — Sample datasets for public release, demonstrating 12-article catalog with stock movements

### Changed
- Public source now includes 24 modules (23 .bas + 1 .frm), up from 22 .bas + 1 .frm

## [1.0.0] — 2026-05-07

### Added
- Full VBA Decision Support System for inventory management
- EOQ (Wilson model) calculations
- ABC classification engine
- CMUP (weighted average unit cost) calculator
- ROP/Safety Stock alerts
- ACID-style transaction management with rollback
- Audit trail logging (append-only)
- Demand forecasting (3/7/14-day moving averages)
- PDF export engine
- Dashboard KPI calculations
- UI theming with dark mode support
- QR code generation
- Receipt verification codes
- Approval workflow system
- Clean-slate build system (build.ps1)
- 5-stage verification suite (verify.ps1)
- VBE control suite (vbe.ps1)
- DSS audit framework (dss-audit.ps1)
- Macro test suite (test-macros.ps1)

### Fixed
- **Attribute VB_Name** syntax error — Build script now strips attributes during ThisWorkbook injection
- **Declaration order** — `Private m_LastSyncTime` moved before Public procedures in mod_SyncBridge
- **UDT coercion** — `CalculateAllForecasts` changed from Function to Sub (UDT arrays cannot be Variant)
- **Exit statement** — `Exit Sub` → `Exit Function` in `CommitTransaction` (mod_StockEntry_Logic)
- **Missing procedure** — Added `InitStatusBar` stub to mod_ThemingEngine
- **Runtime control access** — `Me.txtRefDoc.Value` → `Me.Controls("txtRefDoc").Value` in frmStockEntry
- **Circular dependency** — Broke mod_SharedEnvironment ↔ mod_ExportEngine via Application.Run
- **Hardcoded passwords** — Replaced with `mod_Config.MASTER_PWD` references
- **Sheet protection** — Build script now protects all 24 critical sheets automatically
- **EOQ constants read** — Changed from sheet cell access to VBA constant access
- **Call graph analysis** — Fixed false positives from comments and string literals

### Architecture
- 29 .bas modules + 1 .frm form
- 9,483 lines of VBA code
- Zero external dependencies
- Excel 2010+ compatible
- Offline-first design

### Build System
- Clean-slate rebuild from source files
- Automatic sheet protection during build
- P-code cache corruption prevention
- UTF-8 without BOM encoding
- CRLF line endings

### Audit Results (v1.0.0)
| Category | Pass | Warn | Critical |
|----------|------|------|----------|
| Structural Integrity | 5 | 0 | 0 |
| Security | 4 | 1 | 0 |
| Data Integrity | 2 | 2 | 0 |
| Call Graph | 1 | 1 | 0 |
| Performance & Compliance | 3 | 0 | 0 |
| **Total** | **14** | **3** | **0** |

**Status: PASS WITH WARNINGS**

### Known Warnings
- 1 blank article code (data issue)
- 8 non-numeric stock values (data issue)
- 14 orphan modules (false positives — called from ACCUEIL buttons/ThisWorkbook/forms)

[Unreleased]: https://github.com/kamelmh/lsm-vba-core/compare/v1.4.0...HEAD
[1.4.0]: https://github.com/kamelmh/lsm-vba-core/compare/v1.3.0...v1.4.0
[1.3.0]: https://github.com/kamelmh/lsm-vba-core/compare/v1.0.2...v1.3.0
[1.0.2]: https://github.com/kamelmh/lsm-vba-core/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/kamelmh/lsm-vba-core/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/kamelmh/lsm-vba-core/releases/tag/v1.0.0
