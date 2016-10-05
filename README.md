# excel-macro-vba-library
This repository is created to build a library of reusable VBA / Macro modules and functions for Microsoft Excel.

right click to import the bas file as a new code module in excel vba



included generic funtions and subroutines

WorksheetExists

CopyRowAboveAndInsert

ProperX
'This is exactly like the PROPER function but works also in VBA
'It makes the first letter in the input string capital

Email_part
'Returns the selected part of an email adress

current_user
'gets the user id of the currently logged on active user

current_user_domain
'gets the user domain of the currently logged on active user

get_user_info
'this function queries active directory for information about a user and returns one value at a time

get_user_id_from_distinguishedname
'this function queries active directory for the user id of a user with a certain distinguished name

get_user_id_from_email
'This functions is for when the LDAP query returns the Dist name but we want the ID

ReturnNthPartOfString
'Splits a delimited string and returns the Nth part

RegisterDescriptionForUserDefinedFunction()
' This function is only needed to run once to register help text and category for a user defined function

Custom_GetLastRow

Custom_SetRowHeightSheet
'Adjust row height for the entire sheet.

Custom_CopyRenameSheet
' Sub routine to copy a source sheet to a new sheet and to rename the new sheet.

Custom_NewRenameSheet
' Sub routine to create a new sheet and to rename the new sheet.

Custom_CopyPasteColumn
' Sub routine to copy a column at position 'columnReferenceNumber' and insert it at position 'pastePositionReference'.

Custom_InsertRenameColumn
' Sub routine to insert a new column at column position 'targetColumnReference' of the sheet 'targetSheet', set width as 'targetColumnWidth' and set name as 'targetColumnName'.

Custom_AddComment
' Add a comment to a cell.

Custom_RearrangeColumns
' Sub Routine to re-arrange columns from one sheet to another.  
' A continuous column sequencing and definitions of order of rearrangement in the 'formatReferenceSheetName' is expected for this sub routine to work properly.
' Refer to "FR_1" sheet within the excel file "Excel_Reference_Sheet.xlsx" in the GitHub repository "excel-macro-vba-library".

Custom_EnterFormulaAndFillDown
' This sub routine enters a formula text 'formulaText' in a column's first cell (defined by 'targetSheet', 'columnReference' and 'rowOffset').  
'The formula will be populated down to 'lastRow'.  Furthermore, the function will replace the cells involved with their values and remove the formula definitions after the fill down has been completed.

Custom_ConvertNumberSavedAsText
' Sub routine to convert number stored as text to number.

Custom_DeleteColumn
' Delete the 'targetColumn' in 'targetSheet'.

Custom_CreatePivotTable
' Sub routine to create and design a Pivot Table as per definition in the Reference sheet formatReferenceSheet.
' Refer to "PR_1" and "PR_2" sheet within the excel file "Excel_Reference_Sheet.xlsx" in the GitHub repository "excel-macro-vba-library".

Custom_PivotTableAddField
' Create a Page Field (Report Filter) in the pivot table 'pivotTableName' in sheet 'pivotTableTargetSheet'.

Custom_PivotTableAddDataField
' Create a DataField in the pivot table 'pivotTableName' in sheet 'pivotTableTargetSheet', with name as 'dataFieldName' and with format 'dataFieldFormat'.

Custom_SetColumnNumberFormat
' Set the number format of a particular Column.

Custom_SortSheetByColumn
' Sort the entire sheet 'targetSheet' by the 'key1ColumnReference' in the order indicated by 'order1Reference'

Custom_RemoveDuplicates
' Remove all rows where the Column referred by 'indexColumnReference' has duplicate values.

Custom_FreezeView
' Freeze the view of the targetSheet.  The Split will be made at columnSplitLength and rowSplitLength.

Custom_DeleteSheet
' Delete targetSheet.

Custom_ColorRange
' Enter color into the range.
    
Custom_HideSheet(targetSheet As Worksheet)
' Hide Sheet.

Custom_ColumnFilter
' Enable filter on a column of the targetSheet.  Filter for string value criteriaString.

Custom_ReleaseFilter
' Remove all filters from the targetSheet.

Custom_DeleteVisibleRows
' Delete all Visible Rows in the targetSheet.  This should be used after filtering the current sheet for the information you would want to have deleted.    

Custom_SetColumnWidthSheet
'  Adjust column width for the entire sheet.

Custom_InsertRenameColumn








