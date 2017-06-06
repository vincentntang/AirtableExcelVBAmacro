Sub Woho()

Dim folderPath As String
Dim folderLocation As String
Dim row As Integer
Dim A As String
Dim C As String

    folderPath = Application.ActiveWorkbook.Path
    folderLocation = "hello"


    For row = 1 To 10
        A = Cells(row, 1).Value
        C = Cells(row, 3).Value
        
        A = """" & folderPath & "\" & folderLocation & "\" & A & ".png" & """"
        C = """" & folderPath & "\" & C & """"
        
        Cells(row, 4).Value = "Copy " & C & " " & A
    Next row
    
    

End Sub
