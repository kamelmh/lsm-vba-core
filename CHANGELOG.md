# Changelog

All notable changes to LSM VBA Core will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [Unreleased]

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

[Unreleased]: https://github.com/anomalyco/lsm-vba-core/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/anomalyco/lsm-vba-core/releases/tag/v1.0.0
