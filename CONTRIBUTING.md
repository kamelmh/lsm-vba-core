# Contributing to LSM VBA Core

Thank you for your interest in contributing! This guide will help you get started.

## Development Setup

### Prerequisites
- Microsoft Excel 2010+ (Windows)
- PowerShell 7+
- Git

### Quick Start

```powershell
# 1. Clone the repository
git clone https://github.com/kamelmh/lsm-vba-core.git
cd lsm-vba-core

# 2. Open in Excel
# Create a new workbook, then import all .bas/.frm files via VBE (Alt+F11)

# 3. Or use the build script (requires a master template)
.\tools\build.ps1
```

## Code Conventions

### VBA Style
- **Module names**: PascalCase with `mod_` prefix (e.g., `mod_StockEngine`)
- **Procedure names**: PascalCase (e.g., `CalculateCMUP`, `GetArticleStock`)
- **Variable names**: camelCase or PascalCase
- **Constants**: UPPER_SNAKE_CASE (e.g., `WORKING_DAYS_PER_YEAR`)
- **Comments**: In French, prefixed with `'`

### Critical Rules

1. **No XLOOKUP/FILTER** — Must work in Excel 2010
2. **Option Explicit** — Required in all modules
3. **Error handling** — Use `On Error GoTo` in public procedures
4. **Declaration order** — All `Const`, `Type`, and `Private` declarations before any `Public` procedures
5. **No Attribute VB_Name** in injected code — Document modules have this internally
6. **No comments after line continuation** — `& _ ' comment` breaks VBA

### Common Patterns

```vba
' Good: UDT array via ByRef Sub
Public Sub CalculateAllForecasts(ByRef results() As ForecastResult)
    ' ...
End Sub

' Bad: UDT array as Variant return (compile error)
' Public Function CalculateAllForecasts() As Variant

' Good: Runtime control access
Me.Controls("txtRefDoc").Value = "BR-2026-001"

' Bad: Direct member access for runtime-created controls
' Me.txtRefDoc.Value = "BR-2026-001"
```

## Debug Workflow

When VBA shows an error:

1. Copy highlighted code → save to `Desktop\handoffN.txt`
2. AI reads handoff, diagnoses error type
3. AI fixes source `.bas` file in `VBA_Modules/`
4. Run `build.ps1` to rebuild workbook
5. Run `verify.ps1` to validate
6. Report: Fix applied, Build OK, Safe to open

## Build Protocol

**Never modify the `.xlsm` directly.** Always:

1. Fix source `.bas`/`.frm` files
2. Run `build.ps1`:
   - Kills Excel
   - Opens MASTER workbook
   - Strips ALL user modules
   - Imports fresh source files
   - Compiles
   - Saves as NEW file (forces fresh p-code cache)
3. Run `verify.ps1`
4. Test in Excel

## Testing

```powershell
# Run macro test suite
.\tests\test-macros.ps1

# Run full DSS audit
.\tests\dss-audit.ps1

# Protect sheets (applied automatically during build)
.\tests\protect-sheets.ps1
```

## Audit Checklist

Before submitting changes, ensure:

- [ ] All modules compile without errors
- [ ] No circular dependencies between modules
- [ ] Passwords use `mod_Config.MASTER_PWD` (not hardcoded)
- [ ] Private declarations appear before Public procedures
- [ ] No `Exit Sub` in `Function` procedures
- [ ] No `Attribute VB_Name` in injected document module code
- [ ] DSS audit passes all CRITICAL checks

## Pull Requests

1. Create a feature branch (`git checkout -b feature/my-change`)
2. Make your changes to `.bas`/`.frm` files
3. Run `build.ps1` and `verify.ps1`
4. Commit with a descriptive message
5. Push and create a pull request

## Module Structure

```
src/
├── Core Engine          # Stock calculations, transactions, database
├── Analytics            # Forecasting, dashboards, reports, export
├── UI Framework         # Theming, navigation, localization
├── Utilities            # Shared helpers, audit trail, approval
└── Build Tools          # build.ps1, verify.ps1, vbe.ps1
```

## Known Gotchas

| Issue | Solution |
|-------|----------|
| `Private` declaration after `Public` proc | Move all declarations to top of module |
| UDT array → Variant coercion error | Use `ByRef Sub` pattern instead of `Function` |
| `Exit Sub` in `Function` | Change to `Exit Function` |
| `Attribute VB_Name` syntax error | Strip from source; build.ps1 handles injection |
| `Me.txtRef.Value` not found | Use `Me.Controls("txtRef").Value` for runtime controls |
| Circular dependency detected | Use `Application.Run` for indirect calls |

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
