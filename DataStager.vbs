Sub airtableCleaner()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim x As Integer
    Dim counter As Integer
    Dim argCounter As Integer
    Dim A As String
    Dim B As String
    
    Dim Answer As VbMsgBoxResult
        
    Answer = MsgBox("Do you want to run this macro? Please use airtable Download as CSV - Column 1: Primary key, Column 2: Airtable Link", vbYesNo, "Run Macro")
        
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
    
    'Cleanup Broken links %5B1%5D in Column C
    Columns("C:C").Select
    Range("C40").Activate
    Selection.Replace What:="%5B1%5D", Replacement:="[1]", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    
        
    'Clean Up Each File
    For Each Cell In Range("C2:C" & argCounter)
        If Not dict.exists(CStr(Cell.Value)) Then
            dict.Add Key:=CStr(Cell.Value), Item:=argCounter 'item irrelevant
            counter = 0
        Else
            A = Right(Cell.Value, 4)
            B = Left(Cell.Value, Len(Cell.Value) - 4)
            Cell.Value = B & "[" & counter & "]" & A
            Cell.Value = Cell.Value & "[" & counter & "]"
            counter = counter + 1
        End If
    Next Cell
    Set dict = Nothing 'clear dictionary before exit
    End If
End Sub


