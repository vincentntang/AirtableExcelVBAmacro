Option Explicit

Sub airtableCleaner()
    Dim argCounter As Integer
    Dim folderLocation As Variant
    Dim Answer As VbMsgBoxResult
    Dim myPath As String
    Dim folderPath As String
    
    folderPath = Application.ActiveWorkbook.Path
    myPath = Application.ActiveWorkbook.FullName

    'Ask user if they want to run macro
    Answer = MsgBox("Do you want to run this macro? From airtable, Col 1: primaryKey Col2: one image attachment)", vbYesNo, "Run Macro")
    If Answer = vbYes Then
    
    folderLocation = Application.InputBox("Give a subfolder name for directory. E.G. Batch1")
    
    'Creates new folder based on input
    Dim strDir As String
    strDir = folderPath & "\" & folderLocation
     
    If Dir(strDir, vbDirectory) = "" Then
        MkDir strDir
    Else
        MsgBox "Directory exists."
    End If
    
    
    'Cleanup to just amazons3 dl.airtable links
    Columns("B:B").Select
    Selection.Replace What:="* ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=")", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    'Count Cells
    Range("B2").Activate
    Do
        If ActiveCell.Value = "" Then Exit Do
        ActiveCell.Offset(1, 0).Activate
        argCounter = argCounter + 1

    Loop
    
    'Copy Image Links to new cells to format in Column C
    Columns("B:B").Select
    Selection.Copy
    Columns("C:C").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Clean up links to only have names in Column C
    Selection.Replace What:="https://dl.airtable.com/", Replacement:="", _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
    False, ReplaceFormat:=False
    
    'Cleanup Broken links %5B1%5D in Column C
    Columns("C:C").Select
    Range("C40").Activate
    Selection.Replace What:="%5B1%5D", Replacement:="[1]", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    
    'Create Column D batch files
        Range("D2").Select
    Range("D2").Formula = "=CONCATENATE(""COPY "",CHAR(34),C2,CHAR(34),"" "", CHAR(34), " & _
                      Chr(34) & folderPath & "\" & folderLocation & "\" & Chr(34) & ",A2,"".png"",CHAR(34))"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & argCounter + 1)
    
    'Delete header row 1 information
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    
    'Repaste values back into column D removing formulas
        Columns("D:D").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    End If
End Sub
