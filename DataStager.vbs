Sub airtableCleaner()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim x As Integer
    Dim counter As Integer
    Dim argCounter As Integer
    Dim A As String
    Dim B As String
    
    Dim Answer As VbMsgBoxResult
        
    Answer = MsgBox("Do you want to run this macro? Please use airtable Download as CSV - Column 1: Primary key, Column 2: Airtable Linkz", vbYesNo, "Run Macro")
        
    If Answer = vbYes Then
    counter = 1 'this counts and resets on items that are the same
    
    
    
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
    
    'Create the batch files to name in column C
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
    "=CONCATENATE(""COPY "",CHAR(34),RC[-1],CHAR(34),"" "", CHAR(34), ""c:\doggo\"",RC[-3],"".png"",CHAR(34))"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & argCounter + 1)
    'Range("D2:D" & argCounter+1).Select
    
    'Delete row 1 row information
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    
    
    End If
End Sub


