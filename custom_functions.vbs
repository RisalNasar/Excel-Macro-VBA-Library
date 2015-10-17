' Excel Built-in function to check if a string is present within another string.
' InStr(Start, String1, String2, Compare)
' Start - 'Optional. Numeric expression that sets the starting position for each search.
'  If omitted, search begins at the first character position. The start index is 1-based.
'String1 - Required. String expression being searched.
'String2 - Required. String expression sought.
'Compare - Optional. Specifies the type of string comparison.
'  If Compare is omitted, the Option Compare setting determines the type of comparison.
'https://msdn.microsoft.com/en-us/library/8460tsh1(v=vs.90).aspx

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
