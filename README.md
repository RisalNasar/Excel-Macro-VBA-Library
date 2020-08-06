# Excel Macro VBA Library

This repository is created to build a library of reusable VBA / Macro modules and functions for Microsoft Excel.

Right click to import the bas file as a new code module in excel vba.  

Included generic funtions and subroutines:

```
1) WorksheetExists

2) CopyRowAboveAndInsert

3) ProperX

'This is exactly like the PROPER function but works also in VBA
'It makes the first letter in the input string capital

4) Email_part

'Returns the selected part of an email adress

5) current_user

'gets the user id of the currently logged on active user

6) current_user_domain

'gets the user domain of the currently logged on active user

7) get_user_info

'this function queries active directory for information about a user and returns one value at a time

8) get_user_id_from_distinguishedname

'this function queries active directory for the user id of a user with a certain distinguished name

9) get_user_id_from_email

'This functions is for when the LDAP query returns the Dist name but we want the ID

10) ReturnNthPartOfString

'Splits a delimited string and returns the Nth part

11) RegisterDescriptionForUserDefinedFunction()

' This function is only needed to run once to register help text and category for a user defined function

12) Custom_GetLastRow

13) Custom_SetRowHeightSheet

'Adjust row height for the entire sheet.

14) Custom_CopyRenameSheet

' Sub routine to copy a source sheet to a new sheet and to rename the new sheet.

15) Custom_NewRenameSheet

' Sub routine to create a new sheet and to rename the new sheet.

16) Custom_CopyPasteColumn

' Sub routine to copy a column at position 'columnReferenceNumber' and insert it at position 'pastePositionReference'.

17) Custom_InsertRenameColumn

' Sub routine to insert a new column at column position 'targetColumnReference' of the sheet 'targetSheet', set width as 'targetColumnWidth' and set name as 'targetColumnName'.

18) Custom_AddComment

' Add a comment to a cell.

19) Custom_RearrangeColumns

' Sub Routine to re-arrange columns from one sheet to another.  
' A continuous column sequencing and definitions of order of rearrangement in the 'formatReferenceSheetName' is expected for this sub routine to work properly.
' Refer to "FR_1" sheet within the excel file "Excel_Reference_Sheet.xlsx" in the GitHub repository "excel-macro-vba-library".

20) Custom_EnterFormulaAndFillDown

' This sub routine enters a formula text 'formulaText' in a column's first cell (defined by 'targetSheet', 'columnReference' and 'rowOffset').  
'The formula will be populated down to 'lastRow'.  Furthermore, the function will replace the cells involved with their values and remove the formula definitions after the fill down has been completed.

21) Custom_ConvertNumberSavedAsText

' Sub routine to convert number stored as text to number.

22) Custom_DeleteColumn

' Delete the 'targetColumn' in 'targetSheet'.

23) Custom_CreatePivotTable

' Sub routine to create and design a Pivot Table as per definition in the Reference sheet formatReferenceSheet.
' Refer to "PR_1" and "PR_2" sheet within the excel file "Excel_Reference_Sheet.xlsx" in the GitHub repository "excel-macro-vba-library".

24) Custom_PivotTableAddField

' Create a Page Field (Report Filter) in the pivot table 'pivotTableName' in sheet 'pivotTableTargetSheet'.

25) Custom_PivotTableAddDataField

' Create a DataField in the pivot table 'pivotTableName' in sheet 'pivotTableTargetSheet', with name as 'dataFieldName' and with format 'dataFieldFormat'.

26) Custom_SetColumnNumberFormat

' Set the number format of a particular Column.

27) Custom_SortSheetByColumn

' Sort the entire sheet 'targetSheet' by the 'key1ColumnReference' in the order indicated by 'order1Reference'

28) Custom_RemoveDuplicates

' Remove all rows where the Column referred by 'indexColumnReference' has duplicate values.

29) Custom_FreezeView

' Freeze the view of the targetSheet.  The Split will be made at columnSplitLength and rowSplitLength.

30) Custom_DeleteSheet

' Delete targetSheet.

31) Custom_ColorRange

' Enter color into the range.
    
32) Custom_HideSheet(targetSheet As Worksheet)

' Hide Sheet.

33) Custom_ColumnFilter

' Enable filter on a column of the targetSheet.  Filter for string value criteriaString.

34) Custom_ReleaseFilter

' Remove all filters from the targetSheet.

35) Custom_DeleteVisibleRows

' Delete all Visible Rows in the targetSheet.  This should be used after filtering the current sheet for the information you would want to have deleted.    

36) Custom_SetColumnWidthSheet

'  Adjust column width for the entire sheet.

37) Custom_InsertRenameColumn
```







