Attribute VB_Name = "ModuleX"

'This Excel VBA module can be imported into a xlsm in the vba editor by right clicking modules and chosing import
'To export and upload contributions right click on your module and chose export
'pulled from https://github.com/GreenDiary/excel-macro-vba-library/
'contributors:
'DonJohan - Johan@dgk.se, johan.skoglund@accenture.com
'RisalNasar'

Function WorksheetExists(sName As Variant) As Boolean
    If Not sName = "" Then
        WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
    End If
End Function

Sub CopyRowAboveAndInsert()
    'This macro can be added to a button with an "add new row" text
    'It will copy the row above the button and insert it as a new row
    'clear Clipboard to avoid serious bug
    Application.CutCopyMode = False
    'select row of clicked button
    ActiveCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Select
    'Select the row above, copy and insert
    Selection.Offset(-1, 0).Select
    Selection.EntireRow.Copy
    Selection.EntireRow.Insert
    'clear the selection
    Application.CutCopyMode = False
End Sub

Public Function ProperX(xstr As String) As String
    'This is exactly like the PROPER function but works also in VBA
    'It makes the first letter in the input string capital
   ProperX = UCase(Left(xstr, 1)) & Right(xstr, Len(xstr) - 1)
End Function

Public Function Email_part(xstr As String, part As String) As String
    'Returns the selected part of an email adress
    'Part input can be Fname, Lname, Domain
    'Example 1: =Email_part(C3;"Fname")
    'Example 2: Email_part(mailstring,"Lname")
    Dim FirstDotPos As Integer: FirstDotPos = InStr(1, xstr, ".")
    Dim Apos As Integer: Apos = InStr(1, xstr, "@")
    Dim SecDotPos As Integer: SecDotPos = InStr(Apos, xstr, ".")
    Select Case part
        Case "Fname"
            Email_part = ProperX(Left(xstr, FirstDotPos - 1))
        Case "Lname"
                If FirstDotPos > Apos Then
                    'If no lastname we return the first name
                    Email_part = ProperX(Left(xstr, Apos - 1))
                Else
                    Email_part = ProperX(Mid(xstr, FirstDotPos + 1, Apos - FirstDotPos - 1))
                End If
        Case "Domain"
            Email_part = ProperX(Mid(xstr, Apos + 1, Len(xstr) - Apos))
        Case Else
            Email_part = ProperX("Invalid: Part input can be Fname, Lname or Domain")
    End Select
End Function

Public Function current_user() As String
    'gets the user id of the currently logged on active user
    current_user = Environ("Username")
End Function

Public Function current_user_domain() As String
    'gets the user domain of the currently logged on active user
    current_user_domain = Environ("UserDomain")
End Function

Public Function get_user_info(user_id As String, info_part As String) As String
    'this function queries active directory for information about a user and returns one value at a time
    'example get_user_info(current_user();"mail"
    'info_part can be
    'department, mobile, manager, co, mail, distinguishedName, sAMAccountName
    UserName = user_id
    Set rootDSE = GetObject("LDAP://RootDSE")
    base = "<LDAP://" & rootDSE.get("defaultNamingContext") & ">"
    fltr = "(&(objectClass=user)(objectCategory=Person)" & "(sAMAccountName=" & UserName & "))"
    Attr = info_part
    scope = "subtree"
    Set conn = CreateObject("ADODB.Connection")
    conn.Provider = "ADsDSOObject"
    conn.Open "Active Directory Provider"
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = base & ";" & fltr & ";" & Attr & ";" & scope
    Set rs = cmd.Execute
    If info_part = "manager" Then ' if the part we want is the manager we return the id of the manager instead of dist.name
        get_user_info = get_user_id_from_distinguishedname(rs.Fields(Attr).Value)
    Else
        get_user_info = rs.Fields(Attr).Value
    End If
    rs.Close
    conn.Close
End Function

Public Function get_user_id_from_distinguishedname(DistName As String) As String
    'this function queries active directory for the user id of a user with a certain distinguished name
    Set rootDSE = GetObject("LDAP://RootDSE")
    base = "<LDAP://" & rootDSE.get("defaultNamingContext") & ">"
    fltr = "(&(objectClass=user)(objectCategory=Person)" & "(distinguishedName=" & DistName & "))"
    Attr = "sAMAccountName"
    scope = "subtree"
    Set conn = CreateObject("ADODB.Connection")
    conn.Provider = "ADsDSOObject"
    conn.Open "Active Directory Provider"
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = base & ";" & fltr & ";" & Attr & ";" & scope
    Set rs = cmd.Execute
    get_user_id_from_distinguishedname = rs.Fields(Attr).Value
    rs.Close
    conn.Close
End Function

Public Function get_user_id_from_email(Email As String) As String
    'This functions is for when the LDAP query returns the Dist name but we want the ID
    Set rootDSE = GetObject("LDAP://RootDSE")
    base = "<LDAP://" & rootDSE.get("defaultNamingContext") & ">"
    fltr = "(&(objectClass=user)(objectCategory=Person)" & "(mail=" & Email & "))"
    Attr = "sAMAccountName"
    scope = "subtree"
    Set conn = CreateObject("ADODB.Connection")
    conn.Provider = "ADsDSOObject"
    conn.Open "Active Directory Provider"
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = base & ";" & fltr & ";" & Attr & ";" & scope
    Set rs = cmd.Execute
    get_user_id_from_email = rs.Fields(Attr).Value
    rs.Close
    conn.Close
End Function

Public Function ReturnNthPartOfString(strText As String, Instance As Integer, Delimiter As String) As String
    'Splits a delimited string and returns the Nth part
    'Example: =ReturnNthPartOfString($A2;4;"/")
    Dim strArray() As String
    'Split the string into a collection
    strArray = Split(strText, Delimiter)
    ReturnNthPartOfString = strArray(Instance) 'returns the xth part of the collection
End Function

Sub RegisterDescriptionForUserDefinedFunction()
   ' This function is only needed to run once to register help text and category for a user defined function
   ' I have started to fill it out for a few functions that is usable to have availabe inside of excel and you want to use the function in a cell
   Dim FuncName As String
   Dim FuncDesc As String
   Dim Category As String
   Dim ArgDesc(1 To 3) As String

   'FuncName = "get_user_info"
   'FuncDesc = "Returns an info field about a user from active directory"
   'Category = 7 'Text category
   'ArgDesc(1) = "The user ID"
   'ArgDesc(2) = "Name of the information field you want to return (department, mobile, manager, co, mail)"
   
   FuncName = "ReturnNthPartOfString"
   FuncDesc = "Returns the Nth part of a delimited string"
   Category = 7 'Text category
   ArgDesc(1) = "The delimited text string"
   ArgDesc(2) = "The Nth instance in the string that you want to return"
   ArgDesc(3) = "The Char used as delimiter in the text string"

   Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:=Category, _
      ArgumentDescriptions:=ArgDesc
End Sub

'original functions and subroutines by GreenDiary (now collected in this one file)

Function Custom_GetLastRow(targetSheet As Worksheet) As Long

    Custom_GetLastRow = targetSheet.UsedRange.Rows.count
    
    'Custom_GetLastRow = targetSheet.Columns(columnReference).End(xlDown).Row
    
    ''Ctrl + Shift + End
    'lastRow = sht.Cells(sht.Rows.count, "A").End(xlUp).Row
    '
    ''Using UsedRange
    'sht.UsedRange 'Refresh UsedRange
    'lastRow = sht.UsedRange.Rows(sht.UsedRange.Rows.count).Row
    '
    ''Using Table Range
    'lastRow = sht.ListObjects("Table1").Range.Rows.count
    '
    ''Using Named Range
    'lastRow = sht.Range("MyNamedRange").Rows.count
    '
    ''Ctrl + Shift + Down (Range should be first cell in data set)
    'lastRow = sht.Range("A1").CurrentRegion.Rows.count

End Function

'Adjust row height for the entire sheet.
Sub Custom_SetRowHeightSheet(targetSheet As Worksheet, rowHeightSheet As Integer)

    targetSheet.Cells.RowHeight = rowHeightSheet
    
End Sub


' Sub routine to copy a source sheet to a new sheet and to rename the new sheet.
Sub Custom_CopyRenameSheet(sourceSheet As Worksheet, newSheetName As String)
    
    sourceSheet.Copy _
       after:=ActiveWorkbook.Sheets(Sheets.count)
    ActiveSheet.Name = newSheetName
	
	'Clear the clipboard
    ActiveWorkbook.Application.CutCopyMode = False
    
End Sub

' Sub routine to create a new sheet and to rename the new sheet.
Sub Custom_NewRenameSheet(newSheetName As String)
    
    Sheets.Add.Name = newSheetName
	
	'Move the new sheet to the end of the list of sheets.
    ActiveSheet.Move after:=Worksheets(Worksheets.count)
    
End Sub

' Sub routine to copy a column at position 'columnReferenceNumber' and insert it at position 'pastePositionReference'.
Sub Custom_CopyPasteColumn(sourceSheet As Worksheet, copyColumnReference As Long, destinationSheet As Worksheet, pasteColumnReference As Long)

    sourceSheet.Columns(copyColumnReference).Copy
    destinationSheet.Columns(pasteColumnReference).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    destinationSheet.Columns(pasteColumnReference).ColumnWidth = sourceSheet.Columns(copyColumnReference).ColumnWidth
    
	'https://msdn.microsoft.com/en-us/library/office/ff839476.aspx
	'https://msdn.microsoft.com/en-us/library/office/ff837425.aspx
	'PasteSpecial Paste Options below:
	'xlPasteAll                                - Everything will be pasted.
	'xlPasteAllExceptBorders                   - Everything except borders will be pasted.
	'xlPasteAllMergingConditionalFormats        - Everything will be pasted and conditional formats will be merged.
	'xlPasteAllUsingSourceTheme                 - Everything will be pasted using the source theme.
	'xlPasteColumnWidths                        - Copied column width is pasted.
	'xlPasteComments                            - Comments are pasted.
	'xlPasteFormats                             - Copied source format is pasted.
	'xlPasteFormulas                            - Formulas are pasted.
	'xlPasteFormulasAndNumberFormats            - Formulas and Number formats are pasted.
	'xlPasteValidation                          - Validations are pasted.
	'xlPasteValues                              - Values are pasted.
	'xlPasteValuesAndNumberFormats              - Values and Number formats are pasted.

	'Clear the clipboard
    ActiveWorkbook.Application.CutCopyMode = False

End Sub

' Sub routine to insert a new column at column position 'targetColumnReference' of the sheet 'targetSheet', set width as 'targetColumnWidth' and set name as 'targetColumnName'.
Sub Custom_InsertRenameColumn(targetSheet As Worksheet, targetColumnReference As Long, targetColumnWidth As Integer, _
    targetColumnName As String)

    targetSheet.Columns(targetColumnReference).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    targetSheet.Columns(targetColumnReference).ColumnWidth = targetColumnWidth
    targetSheet.Cells(1, targetColumnReference).Value = targetColumnName

End Sub

' Add a comment to a cell.
Sub Custom_AddComment(targetSheet As Worksheet, targetCellRow As Long, targetCellColumn As Long, commentText As String)

    With targetSheet.Range(targetSheet.Cells(targetCellRow, targetCellColumn), targetSheet.Cells(targetCellRow, targetCellColumn))
        If .Comment Is Nothing Then
			' Handle if no comment exist in the cell already.
            .AddComment.Text Text:=commentText & Chr(10) & ""
        Else
			' Handle if a comment exists in the cell already.
            .Comment.Text Text:=commentText & Chr(10) & ""
        End If
        .Comment.Visible = False
    End With
    
End Sub
'-------------------------------- 

' Sub Routine to re-arrange columns from one sheet to another.  
' A continuous column sequencing and definitions of order of rearrangement in the 'formatReferenceSheetName' 
'  is expected for this sub routine to work properly.
' Refer to "FR_1" sheet within the excel file "Excel_Reference_Sheet.xlsx" in the GitHub repository "excel-macro-vba-library".
Sub Custom_RearrangeColumns(formatReferenceSheet As Worksheet)

    Dim rowHolder As Long
    Dim columnHolder As Long
    Dim count As Long
    Dim numberHolder As Long
    
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    
    Set sourceSheet = Worksheets(formatReferenceSheet.Cells(2, 1).Value)
    Set destinationSheet = Worksheets(formatReferenceSheet.Cells(2, 4).Value)
    
    rowHolder = 2
    columnHolder = 6
    count = 0
    While (IsEmpty(formatReferenceSheet.Cells(rowHolder, columnHolder)) = False)

        numberHolder = formatReferenceSheet.Cells(rowHolder, columnHolder).Value
        count = count + 1

        Call Custom_CopyPasteColumn(sourceSheet, numberHolder, destinationSheet, count)


        rowHolder = rowHolder + 1
    Wend


End Sub

Sub Custom_EnterFormulaAndFillDown(targetSheet As Worksheet, columnReference As Integer, rowOffset As Integer, _
        formulaText As String, lastRow As Long)
' This sub routine enters a formula text 'formulaText' in a column's first cell (defined by 'targetSheet',
'   'columnReference' and 'rowOffset').  The formula will be populated down to 'lastRow'.  Furthermore, the function will replace
'   the cells involved with their values and remove the formula definitions after the fill down has been completed.
    
    
    targetSheet.Range(targetSheet.Cells(rowOffset, columnReference), targetSheet.Cells(rowOffset, columnReference)).Formula = formulaText
    targetSheet.Range(targetSheet.Cells(rowOffset, columnReference), targetSheet.Cells(lastRow, columnReference)).FillDown
    targetSheet.Columns(columnReference).Calculate
    
    
    targetSheet.Range(targetSheet.Cells(rowOffset, columnReference), targetSheet.Cells(lastRow, columnReference)).Copy
    targetSheet.Range(targetSheet.Cells(rowOffset, columnReference), targetSheet.Cells(lastRow, columnReference)).PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ActiveWorkbook.Application.CutCopyMode = False

End Sub

Sub Custom_ConvertNumberSavedAsText(targetSheet As Worksheet, targetColumn As Integer)
' Sub routine to convert number stored as text to number.

    With targetSheet.Columns(targetColumn)
        .NumberFormat = "0"
        .Value = .Value
    End With


End Sub

Sub Custom_DeleteColumn(targetSheet As Worksheet, targetColumn As Integer)
' Delete the 'targetColumn' in 'targetSheet'.
    
    targetSheet.Columns(targetColumn).Delete Shift:=xlToLeft


End Sub


Sub Custom_CreatePivotTable(formatReferenceSheet As Worksheet)
' Sub routine to create and design a Pivot Table as per definition in the Reference sheet formatReferenceSheet.
' Refer to "PR_1" and "PR_2" sheet within the excel file "Excel_Reference_Sheet.xlsx" in the GitHub repository "excel-macro-vba-library".

    Dim sourceDataSheetName As String
    Dim sourceDataSheet As Worksheet
    Dim pivotTableName As String
    Dim pivotTableTargetSheet As Worksheet
    Dim startColumnRange As Integer
    Dim endColumnRange As Integer
    Dim lastRow As Long
    
    
    sourceDataSheetName = formatReferenceSheet.Cells(2, 1).Value
    Set sourceDataSheet = Worksheets(sourceDataSheetName)
    startColumnRange = formatReferenceSheet.Cells(4, 1).Value
    endColumnRange = formatReferenceSheet.Cells(6, 1).Value
    pivotTableName = formatReferenceSheet.Cells(8, 1).Value
    Set pivotTableTargetSheet = Worksheets(formatReferenceSheet.Cells(10, 1).Value)

' Delete any existing Pivot Tables in the pivotTableTargetSheet.
    Dim pivotTableHolder As PivotTable
    For Each pivotTableHolder In pivotTableTargetSheet.PivotTables
        pivotTableHolder.TableRange2.Clear
    Next pivotTableHolder

    lastRow = Custom_GetLastRow(sourceDataSheet)
' Create a Pivot Table with pivotTableName.  The source data should be taken from sourceDataSheet and the pivot table should be created
'   in pivotTableTargetSheet.
'    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        sourceDataSheet.Range("A1").CurrentRegion.Address, Version:=xlPivotTableVersion14).CreatePivotTable _
'        TableDestination:=pivotTableTargetSheet.Range("A1"), TableName:=pivotTableName, DefaultVersion _
'        :=xlPivotTableVersion14

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "'" & sourceDataSheetName & "'!" & sourceDataSheet.Range(sourceDataSheet.Cells(1, startColumnRange), sourceDataSheet.Cells(lastRow, endColumnRange)).Address, _
        Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=pivotTableTargetSheet.Range("A1"), TableName:=pivotTableName, DefaultVersion _
        :=xlPivotTableVersion14


'Default value of ManualUpdate property is False wherein a PivotTable report is recalculated
'   automatically on each change. Turn off automatic updation of Pivot Table during the process of its creation to speed up code.

    pivotTableTargetSheet.PivotTables(pivotTableName).ManualUpdate = True

    Dim rowHolder As Integer
    Dim columnHolder As Integer
    Dim stringHolder As String
    Dim stringHolder2 As String
    Dim stringHolder3 As String
    Dim fieldType As Integer
    
    
    rowHolder = 2
    columnHolder = 2
    fieldType = 1

    While (IsEmpty(formatReferenceSheet.Cells(rowHolder, columnHolder)) = False)

        stringHolder = formatReferenceSheet.Cells(rowHolder, columnHolder).Value
        Call Custom_PivotTableAddField(pivotTableTargetSheet, pivotTableName, stringHolder, fieldType)
        rowHolder = rowHolder + 1
        
    Wend

    rowHolder = 2
    columnHolder = 3
    fieldType = 2

    While (IsEmpty(formatReferenceSheet.Cells(rowHolder, columnHolder)) = False)

        stringHolder = formatReferenceSheet.Cells(rowHolder, columnHolder).Value
        Call Custom_PivotTableAddField(pivotTableTargetSheet, pivotTableName, stringHolder, fieldType)
        rowHolder = rowHolder + 1
        
    Wend
    
    rowHolder = 2
    columnHolder = 4
    fieldType = 3

    While (IsEmpty(formatReferenceSheet.Cells(rowHolder, columnHolder)) = False)

        stringHolder = formatReferenceSheet.Cells(rowHolder, columnHolder).Value
        Call Custom_PivotTableAddField(pivotTableTargetSheet, pivotTableName, stringHolder, fieldType)
        rowHolder = rowHolder + 1
        
    Wend

    rowHolder = 2
    columnHolder = 5

    While (IsEmpty(formatReferenceSheet.Cells(rowHolder, columnHolder)) = False)

        stringHolder = formatReferenceSheet.Cells(rowHolder, columnHolder).Value
        stringHolder2 = formatReferenceSheet.Cells(rowHolder, (columnHolder + 1)).Value
        stringHolder3 = formatReferenceSheet.Cells(rowHolder, (columnHolder + 2)).Value
        Call Custom_PivotTableAddDataField(pivotTableTargetSheet, pivotTableName, stringHolder, stringHolder2, stringHolder3)
        rowHolder = rowHolder + 1
        
    Wend
    
'Default value of ManualUpdate property is False wherein a PivotTable report is recalculated
'   automatically on each change. Turn off automatic update of Pivot Table during the process of its creation to speed up code.

    pivotTableTargetSheet.PivotTables(pivotTableName).ManualUpdate = False


End Sub

Sub Custom_PivotTableAddField(pivotTableTargetSheet As Worksheet, pivotTableName As String, fieldName As String, _
    fieldType As Integer)
' Create a Page Field (Report Filter) in the pivot table 'pivotTableName' in sheet 'pivotTableTargetSheet'.

    Dim pivotTableHolder As PivotTable
    Dim pivotFieldHolder As PivotField
    
    Set pivotTableHolder = pivotTableTargetSheet.PivotTables(pivotTableName)
    Set pivotFieldHolder = pivotTableHolder.PivotFields(fieldName)
    
    Select Case fieldType
        Case 1
            pivotFieldHolder.Orientation = xlPageField
        Case 2
            pivotFieldHolder.Orientation = xlRowField
        Case 3
            pivotFieldHolder.Orientation = xlColumnField
    
    End Select

End Sub

Sub Custom_PivotTableAddDataField(pivotTableTargetSheet As Worksheet, pivotTableName As String, dataFieldName As String, _
    dataFieldFunction As String, dataFieldFormat As String)
' Create a DataField in the pivot table 'pivotTableName' in sheet 'pivotTableTargetSheet', with name as 'dataFieldName' and with
'   format 'dataFieldFormat'.

    Dim pivotTableHolder As PivotTable
    Dim functionType As Integer
    
    Select Case dataFieldFunction
        Case "Sum"
            functionType = xlSum
        Case "Count"
            functionType = xlCount
        Case "Maximum"
            functionType = xlMax
        Case "Minimum"
            functionType = xlMin
    
    End Select

    Set pivotTableHolder = pivotTableTargetSheet.PivotTables(pivotTableName)
    
    With pivotTableHolder.PivotFields(dataFieldName)
        .Orientation = xlDataField
        .Function = functionType
        .NumberFormat = dataFieldFormat
    End With


' https://msdn.microsoft.com/en-us/library/office/ff837374.aspx
' XlConsolidationFunction Enumeration
'Name               -   Description
'xlAverage          -   Average.
'xlCount            -   Count.
'xlCountNums        -   Count numerical values only.
'xlDistinctCount    -   Count using Distinct Count analysis.
'xlMax              -   Maximum.
'xlMin              -   Minimum.
'xlProduct          -   Multiply.
'xlStDev            -   Standard deviation, based on a sample.
'xlStDevP           -   Standard deviation, based on the whole population.
'xlSum              -   Sum.
'xlUnknown          -   No subtotal function specified.
'xlVar              -   Variation, based on a sample.
'xlVarP             -   Variation, based on the whole population.


End Sub

Sub Custom_SetColumnNumberFormat(targetSheet As Worksheet, columnReference As Integer, numberFormatString As String)
' Set the number format of a particular Column.

    targetSheet.Columns(columnReference).NumberFormat = numberFormatString

	'References for numberFormatString:
	'https://msdn.microsoft.com/en-us/library/office/ff196401.aspx
	'https://support.office.com/en-us/article/Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-US&rs=en-US&ad=US
	
End Sub

Sub Custom_SortSheetByColumn(targetSheet As Worksheet, key1ColumnReference As Integer, order1String As String)
' Sort the entire sheet 'targetSheet' by the 'key1ColumnReference' in the order indicated by 'order1Reference'

    Dim orderChoice As Integer

    Select Case order1String
        Case "Ascending"
            orderChoice = xlAscending
        Case "Descending"
            orderChoice = xlDescending
    End Select

    targetSheet.Range("A1").CurrentRegion.Sort key1:=targetSheet.Columns(key1ColumnReference), order1:=orderChoice, Header:=xlYes

End Sub

Sub Custom_RemoveDuplicates(targetSheet As Worksheet, indexColumnReference As Integer)
' Remove all rows where the Column referred by 'indexColumnReference' has duplicate values.

    targetSheet.Range("A1").CurrentRegion.RemoveDuplicates Columns:=indexColumnReference, Header:=xlYes

End Sub

Sub Custom_FreezeView(targetSheet As Worksheet, columnSplitLength As Long, rowSplitLength As Long)
' Freeze the view of the targetSheet.  The Split will be made at columnSplitLength and rowSplitLength.

    targetSheet.Activate

    With ActiveWindow
        .SplitColumn = columnSplitLength
        .SplitRow = rowSplitLength
    End With
    ActiveWindow.FreezePanes = True

End Sub

Sub Custom_DeleteSheet(targetSheet As Worksheet)
' Delete targetSheet.
    
    'Stopping Application Alerts
    ActiveWorkbook.Application.DisplayAlerts = False
    
    targetSheet.Delete
    
    'Enabling Application alerts once we are done with our task
    ActiveWorkbook.Application.DisplayAlerts = True

End Sub

Sub Custom_ColorRange(targetSheet As Worksheet, rowStartCoordinate As Long, columnStartCoordinate As Long, _
    rowEndCoordinate As Long, columnEndCoordinate As Long, colorRedValue As Integer, colorGreenValue As Integer, _
    colorBlueValue As Integer)
    ' Enter color into the range.
    
    targetSheet.Range(targetSheet.Cells(rowStartCoordinate, columnStartCoordinate), _
        targetSheet.Cells(rowEndCoordinate, columnEndCoordinate)).Interior.Color = RGB(colorRedValue, colorGreenValue, colorBlueValue)
        
    
    End Sub
    
Sub Custom_HideSheet(targetSheet As Worksheet)
' Hide Sheet.

    targetSheet.Visible = xlSheetHidden

End Sub

Sub Custom_ColumnFilter(targetSheet As Worksheet, columnReference As Integer, criteriaString As String)
' Enable filter on a column of the targetSheet.  Filter for string value criteriaString.

    targetSheet.Columns(columnReference).AutoFilter field:=1, Criteria1:=criteriaString, VisibleDropDown:=False
    

End Sub

Sub Custom_ReleaseFilter(targetSheet As Worksheet)
' Remove all filters from the targetSheet.
    
    targetSheet.AutoFilterMode = False
    
End Sub

Sub Custom_DeleteVisibleRows(targetSheet As Worksheet)
' Delete all Visible Rows in the targetSheet.  This should be used after filtering the current sheet for the information
'	you would want to have deleted.    
	
    targetSheet.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete

End Sub

Sub Custom_SetColumnWidthSheet(targetSheet As Worksheet, columnWidthSheet As Integer)
 '  Adjust column width for the entire sheet.
 
    targetSheet.Cells.ColumnWidth = columnWidthSheet
    
End Sub

Sub Custom_InsertRenameColumn(targetSheet As Worksheet, targetColumnReference As Integer, targetColumnWidth As Integer, _
    targetColumnName As String)
' Insert a new column at the tagetColumnReference of the targetSheet, set width as targetColumnWidth and set name as targetColumnName

    targetSheet.Columns(targetColumnReference).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    targetSheet.Columns(targetColumnReference).ColumnWidth = targetColumnWidth
    targetSheet.Cells(1, targetColumnReference).Value = targetColumnName

End Sub



