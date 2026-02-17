# API Reference

Complete reference for all procedures in the Excel Macro VBA Library. Parameters are described where useful.

---

## Worksheet helpers

### `WorksheetExists(sName As Variant) As Boolean`

Returns whether a sheet with the given name exists in the workbook.  
Example: `If WorksheetExists("Data") Then ...`

---

### `Custom_GetLastRow(targetSheet As Worksheet) As Long`

Returns the last used row on `targetSheet` (based on `UsedRange`).  
Example: `lastRow = Custom_GetLastRow(ActiveSheet)`

---

### `Custom_CopyRenameSheet(sourceSheet As Worksheet, newSheetName As String)`

Copies `sourceSheet` to a new sheet at the end of the workbook and renames it to `newSheetName`.

---

### `Custom_NewRenameSheet(newSheetName As String)`

Creates a new sheet, renames it to `newSheetName`, and moves it to the end of the sheet tabs.

---

### `Custom_DeleteSheet(targetSheet As Worksheet)`

Deletes `targetSheet`. Uses `DisplayAlerts = False` to suppress the confirmation dialog.

---

### `Custom_HideSheet(targetSheet As Worksheet)`

Hides `targetSheet` (xlSheetHidden).

---

### `Custom_SetRowHeightSheet(targetSheet As Worksheet, rowHeightSheet As Integer)`

Sets the row height for all rows on `targetSheet` to `rowHeightSheet`.

---

### `Custom_SetColumnWidthSheet(targetSheet As Worksheet, columnWidthSheet As Integer)`

Sets the column width for all columns on `targetSheet` to `columnWidthSheet`.

---

## Strings and text

### `ProperX(xstr As String) As String`

Same idea as Excel’s PROPER: capitalizes the first letter of `xstr`. Works in VBA (unlike the worksheet PROPER in some contexts).  
Example: `ProperX("hello")` → `"Hello"`

---

### `Email_part(xstr As String, part As String) As String`

Extracts part of an email address.  
**part:** `"Fname"` (before first dot), `"Lname"` (between first dot and @), `"Domain"` (after @).  
Example: `=Email_part(C3,"Domain")`

---

### `ReturnNthPartOfString(strText As String, Instance As Integer, Delimiter As String) As String`

Splits `strText` by `Delimiter` and returns the Nth part (1-based).  
Example: `=ReturnNthPartOfString(A2, 4, "/")`

---

## User and Active Directory

*Require a domain-joined machine and AD access.*

### `current_user() As String`

Returns the Windows username of the current user (`Environ("Username")`).

---

### `current_user_domain() As String`

Returns the Windows user domain (`Environ("UserDomain")`).

---

### `get_user_info(user_id As String, info_part As String) As String`

Returns one attribute from Active Directory for `user_id`.  
**info_part:** e.g. `"department"`, `"mobile"`, `"manager"`, `"co"`, `"mail"`, `"distinguishedName"`, `"sAMAccountName"`.  
Example: `get_user_info(current_user(), "mail")`

---

### `get_user_id_from_distinguishedname(DistName As String) As String`

Returns the sAMAccountName (user id) for the user with the given LDAP distinguished name.

---

### `get_user_id_from_email(Email As String) As String`

Returns the sAMAccountName for the user with the given `mail` attribute in AD.

---

## Rows and columns

### `CopyRowAboveAndInsert()`

Designed to be run from a button: copies the row above the button’s cell and inserts it below. Assumes the button’s `TopLeftCell` is the active context. Clears clipboard when done.

---

### `Custom_CopyPasteColumn(sourceSheet, copyColumnReference, destinationSheet, pasteColumnReference)`

Copies one column from `sourceSheet` and pastes values and number formats into `destinationSheet` at the given column index. Column width is copied.  
Parameters: `copyColumnReference`, `pasteColumnReference` are `Long` column indices.

---

### `Custom_InsertRenameColumn(targetSheet, targetColumnReference, targetColumnWidth, targetColumnName)`

Inserts a new column at `targetColumnReference`, sets its width and sets the header cell (row 1) to `targetColumnName`.  
Overload in module: one version uses `Long` for column, one uses `Integer`; behavior is the same.

---

### `Custom_DeleteColumn(targetSheet As Worksheet, targetColumn As Integer)`

Deletes the column at `targetColumn` on `targetSheet` (shifts cells left).

---

### `Custom_RearrangeColumns(formatReferenceSheet As Worksheet)`

Rearranges columns from a source sheet to a destination sheet using a reference sheet that defines order.  
Expects continuous column indices and layout as in the “FR_1” sheet of `Excel_Reference_Sheet.xlsx` (see repo reference).

---

## Cells and formatting

### `Custom_AddComment(targetSheet, targetCellRow, targetCellColumn, commentText)`

Adds or replaces the comment on the cell at the given row/column with `commentText`. Hides the comment after setting.

---

### `Custom_ColorRange(targetSheet, rowStartCoordinate, columnStartCoordinate, rowEndCoordinate, columnEndCoordinate, colorRedValue, colorGreenValue, colorBlueValue)`

Fills the range from (rowStart, colStart) to (rowEnd, colEnd) with the RGB color given.

---

### `Custom_SetColumnNumberFormat(targetSheet, columnReference, numberFormatString)`

Sets the number format of the entire column `columnReference` to `numberFormatString` (e.g. `"0"`, `"0.00"`, `"dd/mm/yyyy"`).

---

### `Custom_ConvertNumberSavedAsText(targetSheet, targetColumn)`

Converts numbers stored as text in `targetColumn` to numeric values (format "0" + assign value back).

---

## Formulas and data

### `Custom_EnterFormulaAndFillDown(targetSheet, columnReference, rowOffset, formulaText, lastRow)`

Puts `formulaText` in the first cell of the column (at `rowOffset`), fills down to `lastRow`, then converts the range to values (removes formulas).

---

### `Custom_SortSheetByColumn(targetSheet, key1ColumnReference, order1String)`

Sorts the current region of `targetSheet` by the column `key1ColumnReference`.  
**order1String:** `"Ascending"` or `"Descending"`.

---

### `Custom_RemoveDuplicates(targetSheet, indexColumnReference)`

Removes duplicate rows based on the column at `indexColumnReference`. Header row is assumed (Header:=xlYes).

---

### `Custom_ColumnFilter(targetSheet, columnReference, criteriaString)`

Applies AutoFilter on the used range; filters the column at `columnReference` by `criteriaString`.  
Note: implementation uses `field:=1` on `Columns(columnReference)`; for multi-column ranges the column index within that range may need to match intent.

---

### `Custom_ReleaseFilter(targetSheet As Worksheet)`

Removes AutoFilter from `targetSheet`.

---

### `Custom_DeleteVisibleRows(targetSheet As Worksheet)`

Deletes all visible rows in the used range (offset by one row). Typically used after applying a filter to remove the visible (filtered-in) rows.

---

## View

### `Custom_FreezeView(targetSheet, columnSplitLength, rowSplitLength)`

Freezes panes on `targetSheet`: split at `columnSplitLength` columns and `rowSplitLength` rows, then sets `FreezePanes = True`.

---

## Pivot tables

*Reference layout: “PR_1” and “PR_2” in `Excel_Reference_Sheet.xlsx`.*

### `Custom_CreatePivotTable(formatReferenceSheet As Worksheet)`

Creates a pivot table from a reference sheet that defines source sheet, column range, pivot name, target sheet, and field layout (page/row/column/data fields).

---

### `Custom_PivotTableAddField(pivotTableTargetSheet, pivotTableName, fieldName, fieldType)`

Adds a field to the pivot table.  
**fieldType:** 1 = Page (Report Filter), 2 = Row, 3 = Column.

---

### `Custom_PivotTableAddDataField(pivotTableTargetSheet, pivotTableName, dataFieldName, dataFieldFunction, dataFieldFormat)`

Adds a data field.  
**dataFieldFunction:** `"Sum"`, `"Count"`, `"Maximum"`, `"Minimum"`.  
**dataFieldFormat:** any valid number format string.

---

## UDF registration

### `RegisterDescriptionForUserDefinedFunction()`

Registers help text and category for user-defined functions so they appear in the Insert Function dialog. Run once per session (or from Auto_Open). The code includes an example for `ReturnNthPartOfString`; uncomment and duplicate blocks for other UDFs.

---

## UTF-8 file I/O (`Lib_ReadWriteUtf8.vb`)

### `ReadFromUTFFile(filepath As String) As String`

Reads the entire file at `filepath` as UTF-8 and returns its contents as a string.

---

### `WriteToUTFFile(filepath As String, content As String)`

Writes `content` to the file at `filepath` using UTF-8 encoding (overwrites if the file exists).
