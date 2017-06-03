Sub airtableCleaner()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim x As Integer
    Dim argCounter As Integer
    Dim A As String
    Dim B As String
    Dim folderLocation As Variant
    Dim Answer As VbMsgBoxResult
        
    'Ask user if they want to run macro
    Answer = MsgBox("Do you want to run this macro? Please use airtable Download as CSV - Column 1: Primary key, Column 2: Airtable Linkz", vbYesNo, "Run Macro")
    If Answer = vbYes Then
    
    folderLocation = Application.InputBox("C:\doge\ Enter a folder location where your image assets will be")
    
    
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

    
    'Create Batch File Maker Column D
        Range("D2").Select
    Range("D2").Formula = "=CONCATENATE(""COPY "",CHAR(34),C2,CHAR(34),"" "", CHAR(34), " & _
                      Chr(34) & folderLocation & Chr(34) & ",A2,"".png"",CHAR(34))"
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


