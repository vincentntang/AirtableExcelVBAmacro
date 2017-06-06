Sub Woho()

Dim folderPath As String
Dim folderLocation As String
Dim row As Integer
Dim a As String

'http://i.imgur.com/AbKT80k.png how this path folder works

    folderPath = Application.ActiveWorkbook.Path
    folderLocation = "hello"

 Range("D2").Select
    

    For row = 1 To 10
        a = Cells(row, 1).Value
        B = Cells(row, 3).Value
        
        Cells(row, 4).Value = "Copy " & a & " " & B
    Next row
    
   
End Sub
