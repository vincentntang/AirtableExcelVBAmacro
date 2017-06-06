Sub Woho()

Dim folderPath As String
Dim folderLocation As String
Dim row As Integer

'http://i.imgur.com/AbKT80k.png how this path folder works

    folderPath = Application.ActiveWorkbook.Path
    folderLocation = "hello"

 Range("D2").Select
    '    "=CONCATENATE(""COPY "",CHAR(34), " & folderPath & ",C2,CHAR(34),"" "", CHAR(34), " & _

    For row = 1 To 6
        A = A + 1
    Next row
    Range("D2").Value = A


    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & 10)

End Sub

