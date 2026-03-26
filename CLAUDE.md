# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel VBA monorepo containing structural engineering calculation tools and administrative utilities. Each subdirectory (named by date, e.g. `20170524_RCScan/`) is an independent Excel-based tool (`.xlsm`) with its VBA source exported as `.vb`/`.bas`/`.cls` files. Comments and UI are in Traditional Chinese.

## Development Workflow (SOP)

1. Edit VBA source files (`.vb`, `.bas`, `.cls`)
2. Import code into Excel workbook: `Ctrl + Shift + I`
3. Export code from Excel workbook: `Ctrl + Shift + E`
4. Update Release Notes sheet with version info
5. Update `version.bas` / `Version.txt` with new version number
6. Commit to git
7. Deploy updated `VERSION.json` and `.xlsm` to AWS

Import/Export is handled by `utils/ImportAndExportVBA.bas` which requires the VBIDE extensibility library.

## Architecture

### Shared Utilities (`utils/`)

- **UTILS_CLASS.vb** — Core utility class used across projects. Provides: `CreateDictionary()`, `GetRangeToArray()`, `ExecutionTime()`, `QuickSortArray()`, `ParseJSON()`, `OpenTextFile()`, math helpers (`RoundUp`, `Min`, `Max`)
- **VERSION.vb** — Version checking system that fetches latest version from GitHub and prompts users to update. Uses cloud-based password validation via `VerifyPassword()`
- **ImportAndExportVBA.bas** — Synchronizes VBA code between `.xlsm` workbooks and source files on disk

### Per-Project Structure

Each project follows a consistent pattern:
```
20YYMMDD_ProjectName/
├── ProjectName.xlsm     # Main Excel workbook
├── MainSub.vb           # Entry point subroutines
├── *Class.vb            # Domain classes (e.g. BeamClass, ColumnClass)
├── version.vb           # Version management (imports shared VERSION.vb pattern)
├── Version.txt          # Current version string (semver)
└── __test__/            # Test data workbooks
```

Projects use class-based OOP architecture. Structural engineering projects typically have domain-specific classes (beams, columns, slabs) with calculation logic.

## File Encoding

VB files (`.vb`) use **CP950** encoding (Traditional Chinese Big5), as configured in `.vscode/settings.json`. Ensure correct encoding when reading/writing these files.
