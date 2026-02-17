Function ReadFromUTFFile(filepath As String) As String
'Read text from file encoded with UTF-8
    Dim objStream As Object
    Dim rawfile As String
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (filepath)
    rawfile = ""
    'Datei auslesen
    rawfile = objStream.ReadText
    objStream.Close
    Set objStream = Nothing
    ReadFromUTFFile = rawfile
End Function
Function WriteToUTFFile(filepath As String, content As String)
'Write text to file encoded with UTF-8
    Dim FileName As String
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.WriteText (content)
    objStream.SaveToFile filepath, 2
    objStream.Close
    Set objStream = Nothing
End Function