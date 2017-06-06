Sub Woho()

Dim folderPath As String
Dim folderLocation As String

    folderPath = Application.ActiveWorkbook.Path
    folderLocation = "hello"

 Range("D2").Select
    
    Range("D2").Formula = _
    "=CONCATENATE(""COPY "",CHAR(34),C2,CHAR(34),"" "", CHAR(34), " & _
                      Chr(34) & folderPath & "\" & folderLocation & "\" & Chr(34) & _
                      ",A2,"".png"",CHAR(34))"
                      
   

    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & 10)

End Sub
