
Sub Custom_ColumnFilter(targetSheet As Worksheet, columnReference As Integer, criteriaString As String)

    targetSheet.Columns(columnReference).AutoFilter field:=1, Criteria1:=criteriaString, VisibleDropDown:=False
    

End Sub

Sub Custom_ReleaseFilter(targetSheet As Worksheet)
    
    targetSheet.AutoFilterMode = False
    
End Sub

Sub Custom_DeleteVisibleRows(targetSheet As Worksheet)
    
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
