# Excel Macro VBA Library

> **Note:** This repository is not actively maintained. The code is provided as-is for use and forking; issues and PRs may not get a timely response.

A library of reusable VBA modules and functions for Microsoft Excel. Use it to speed up automation, user-defined functions (UDFs), and common worksheet operations.

## Table of contents

- [Requirements](#requirements)
- [Installation](#installation)
- [Quick start](#quick-start)
- [What's included](#whats-included)
- [API reference](#api-reference)
- [Contributing](#contributing)

## Requirements

- **Microsoft Excel** with macros enabled (`.xlsm` or macro-enabled workbook).
- **Active Directory**: Functions `get_user_info`, `get_user_id_from_distinguishedname`, and `get_user_id_from_email` require a domain-joined machine and AD access.

## Installation

1. Open your workbook in Excel and press **Alt+F11** to open the VBA editor.
2. In the **Project** pane, right-click your project (or the **Modules** folder).
3. Choose **Import File…** and select the `.bas` or `.vb` file you want (e.g. `Module1.bas`).
4. Save the workbook as a **Macro-Enabled Workbook** (`.xlsm`).

To **export** a module (e.g. to contribute): right-click the module in the VBA editor → **Export File**.

## Quick start

- **As UDF in a cell:** e.g. `=ProperX(A1)`, `=Email_part(B2,"Domain")`, `=ReturnNthPartOfString(A2,4,"/")`.
- **As VBA:** call any `Public` function or sub from your own macros, e.g. `Custom_GetLastRow(ActiveSheet)`.

Run `RegisterDescriptionForUserDefinedFunction` once per session (or from an Auto_Open macro) to register help text and category for UDFs so they appear in the Insert Function dialog.

## What's included

| Category | Description |
|----------|-------------|
| **Worksheets** | Check if a sheet exists, copy/rename sheets, get last row, set row/column dimensions. |
| **Strings & text** | `ProperX` (PROPER in VBA), `Email_part`, `ReturnNthPartOfString`. |
| **User & AD** | Current user/domain, and Active Directory lookups (manager, mail, department, etc.). |
| **Columns & rows** | Copy/paste/insert/rename/delete columns, filters, sort, remove duplicates, freeze panes. |
| **Cells & formatting** | Comments, number format, color range, convert numbers stored as text. |
| **Pivot tables** | Create pivot from a reference sheet, add page/row/column/data fields. |
| **Utilities** | Copy row above and insert (button-friendly), UTF-8 file read/write. |

Full list of procedures: [**docs/API.md**](docs/API.md).

## API reference

Detailed signatures, parameters, and usage for every function and sub are in:

- **[docs/API.md](docs/API.md)** — Complete API reference

## File overview

| File | Purpose |
|------|---------|
| `Module1.bas` | Main library: worksheets, strings, AD, columns, pivots, formatting, etc. |
| `Lib_ReadWriteUtf8.vb` | UTF-8 text file read/write for VBA. |
| `custom_modules.vbs` / `custom_functions.vbs` | Reference/snippets (overlap with `Module1.bas`). |

## Contributing

See **[CONTRIBUTING.md](CONTRIBUTING.md)** for how to add or change modules and keep the library consistent.

## Credits

- Contributors: RisalNasar, DonJohan
