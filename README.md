# LSM VBA Core — Logistics & Stock Management Framework

> Pure VBA Decision Support System for inventory management. Zero external dependencies. Excel 2010+ compatible.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![VBA](https://img.shields.io/badge/VBA-Excel%202010%2B-green.svg)](https://docs.microsoft.com/en-us/office/vba/)
[![CI Build](https://github.com/kamelmh/lsm-vba-core/actions/workflows/build.yml/badge.svg)](https://github.com/kamelmh/lsm-vba-core/actions/workflows/build.yml)

## Overview

LSM VBA Core is a **generic, reusable framework** for building inventory management and logistics systems entirely within Microsoft Excel using VBA. No databases, no web servers, no external libraries — just Excel.

### Key Features

- **Offline-First**: Zero network dependencies, all processing local to VBA
- **Decision Support**: EOQ (Wilson), ABC classification, CMUP, ROP/SS calculations
- **Transaction Safety**: ACID-style commit/rollback for stock movements
- **Audit Trail**: Append-only logging for compliance
- **Forecasting**: 3/7/14-day moving average demand prediction
- **Reporting**: PDF export, dashboards, KPI tracking
- **Theming**: Customizable UI with dark mode support
- **Excel 2010 Compatible**: No XLOOKUP, FILTER, or modern functions

## Architecture

```
src/
├── Core Engine
│   ├── mod_Config.bas           # Configuration constants
│   ├── mod_StockEngine.bas      # EOQ, ABC, CMUP, ROP calculations
│   ├── mod_StockEntry_Logic.bas # Form state & transaction commit
│   ├── mod_Database.bas         # Secure write layer
│   └── mod_TransactionSafety.bas# ACID transaction management
├── Analytics
│   ├── mod_SyncBridge.bas       # Data sync & metrics
│   ├── mod_Forecasting.bas      # Demand forecasting
│   ├── mod_Dashboard.bas        # KPI calculation
│   ├── mod_Reports.bas          # Report generation
│   └── mod_ExportEngine.bas     # PDF export
├── UI Framework
│   ├── mod_ThemingEngine.bas    # UI theming
│   ├── mod_UIEnhancements.bas   # Hover effects, animations
│   ├── mod_Navigation.bas       # Sheet navigation
│   ├── mod_Localization.bas     # Multi-language support
│   └── frmStockEntry.frm        # Stock entry form
├── Utilities
│   ├── mod_Utilities.bas        # Shared helpers
│   ├── mod_SheetSetup.bas       # Sheet creation & protection
│   ├── mod_SharedEnvironment.bas# Session management
│   ├── mod_AuditTrail.bas       # Audit logging
│   ├── mod_ApprovalWorkflow.bas # Approval system
│   ├── mod_QRCode.bas           # QR code generation
│   ├── mod_ReceiptTag.bas       # Receipt verification
│   ├── mod_CSVImportExport.bas  # CSV import/export engine
│   └── mod_Barcode.bas          # Keyboard-wedge barcode scanning
└── Build Tools
    ├── build.ps1                # Clean-slate rebuild script
    ├── verify.ps1               # 5-stage verification suite
    └── vbe.ps1                  # VBE control suite
```

## Quick Start

### Prerequisites
- Microsoft Excel 2010 or later (Windows)
- PowerShell 7+ (for build scripts)
- Macro security set to "Enable all macros" (for development)

### Build from Source

```powershell
# Clone the repository
git clone https://github.com/kamelmh/lsm-vba-core.git
cd lsm-vba-core

# Run the build script (creates a new .xlsm from source)
.\tools\build.ps1

# Run verification
.\tools\verify.ps1
```

### Manual Setup

1. Create a new Excel workbook
2. Create the required sheets (ACCUEIL, ARTICLES, MOUVEMENTS, etc.)
3. Import all `.bas` and `.frm` files via VBE (Alt+F11 → File → Import)
4. Set references: Tools → References → check "Microsoft Scripting Runtime"
5. Compile: Debug → Compile VBAProject

## Core Concepts

### EOQ (Economic Order Quantity)
Uses the Wilson formula: `Q* = √(2DS/H)` where D=annual demand, S=order cost, H=holding cost.

### ABC Classification
Pareto-based categorization:
- **A**: Top 70% of value (~20% of items)
- **B**: Next 20% of value (~30% of items)
- **C**: Last 10% of value (~50% of items)

### CMUP
Weighted Average Unit Cost recalculated on every IN transaction.

### ROP / Safety Stock
Reorder Point = (Avg Daily Demand × Lead Time) + Safety Stock

## Development

### Debug Workflow
When VBA shows an error:
1. Save highlighted code to `Desktop\handoffN.txt`
2. AI reads, diagnoses, fixes source `.bas` file
3. Run `build.ps1` to rebuild workbook
4. Run `verify.ps1` to validate

### Clean-Slate Build Protocol
**Never modify the `.xlsm` directly.** Always:
1. Fix source `.bas`/`.frm` files
2. Run `build.ps1` (kills Excel → strips modules → reimports → compiles)
3. Open the new workbook

### Common Pitfalls
- ❌ Private declarations after Public procedures
- ❌ Returning UDT arrays as Variant (use `ByRef Sub` pattern)
- ❌ Direct control access for runtime-created controls (use `Controls(name)`)
- ❌ Attribute VB_Name in injected document module code
- ❌ XLOOKUP/FILTER (Excel 2010 compatibility)

## Testing

```powershell
# Run macro test suite
.\tests\test-macros.ps1

# Run full DSS audit
.\tests\dss-audit.ps1
```

## Audit Results (v1.0.0)

| Category | Status |
|----------|--------|
| Structural Integrity | ✅ 5/5 |
| Security | ⚠️ 3/5 (sheet protection, password exposure, stock types) |
| Data Integrity | ⚠️ 3/4 (blank codes, type validation) |
| Call Graph | ⚠️ 1/2 (orphan modules, circular dep) |
| Performance & Compliance | ✅ 2/3 |

See `tests/dss-audit-report.csv` for details.

## License

MIT License — see [LICENSE](LICENSE) for details.

## Credits

Originally developed for the Algerian Ministry of Education's inventory management system.
Generic core extracted and open-sourced for community use.

Author: Mahi Kamel Abdelghani | CNEPD 2026
