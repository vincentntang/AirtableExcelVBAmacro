Option Explicit

' https://stackoverflow.com/questions/31359682/with-excel-vba-save-web-image-to-disk/31360105#31360105, URL download to file
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
      Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As LongPtr, _
        ByVal lpfnCB As LongPtr _
      ) As Long
    Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "Wininet.dll" _
      Alias "DeleteUrlCacheEntryA" ( _
        ByVal lpszUrlName As String _
      ) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" _
      Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As Long _
      ) As Long
    Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" _
      Alias "DeleteUrlCacheEntryA" ( _
        ByVal lpszUrlName As String _
      ) As Long
#End If

Public Const ERROR_SUCCESS As Long = 0
Public Const BINDF_GETNEWESTVERSION As Long = &H10
Public Const INTERNET_FLAG_RELOAD As Long = &H80000000
                
                ' End of stackoverflow code snippet reference for URLdownloadToFile

    'Global Variables for passing values b/w subs
    Dim myPath As String
    Dim folderPath As String
    Dim folderLocation As Variant 'not used


Sub airtableCleaner()
    Dim argCounter As Integer
    Dim Answer As VbMsgBoxResult
    
    Dim shellCommand As String
    
    Dim row As Integer
    Dim A As String
    Dim C As String


    folderPath = Application.ActiveWorkbook.Path 'Example C:/downloads
    myPath = Application.ActiveWorkbook.FullName 'Example C:/downloads/book1.csv - not used

    'Ask user if they want to run macro
    Answer = MsgBox("Run? Airtable - 1: primaryKey, 2: one image attachment", vbYesNo, "Run Macro")
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
    
    'Cleanup Broken images using excelVBA downloader %5B1%5D = B1D
     Columns("C:C").Select
     Range("C40").Activate
     Selection.Replace What:="%5B1%5D", Replacement:="%%5B1%%5D", LookAt:=xlPart, _
     SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
     ReplaceFormat:=False

    'Create Column D batch files
    For row = 2 To argCounter + 1
        A = Cells(row, 1).Value
        C = Cells(row, 3).Value
        
        A = """" & folderPath & "\" & folderLocation & "\" & A & ".png" & """"
        C = """" & folderPath & "\" & C & """"
        
        Cells(row, 4).Value = "Copy " & C & " " & A
    Next row
    
    'Delete header row 1 information
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    
    'Repaste values back into column D removing formulas
        Columns("D:D").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Image downloader to source folder
    Call dlStaplesImages
    
    'Make the batch files using row data col D
    Call ExportRangetoBatch
    
    'Ask user to run bat file now or later
    shellCommand = """" & folderPath & "\" & "newcurl.bat" & """"
    Call Shell(shellCommand, vbNormalFocus)
    
    End If
End Sub

'https://superuser.com/questions/1045707/create-bat-file-with-excel-data-with-vba ,For exporting column data to batch file to run in shell , modified copypasta code

Sub ExportRangetoBatch()

    Dim ColumnNum: ColumnNum = 4   ' Column D
    Dim RowNum: RowNum = 1         ' Row to start on
    Dim objFSO, objFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(folderPath & "\newcurl.bat")    'Output Path

    Dim OutputString: OutputString = ""

    OutputString = "Timeout 3" & vbNewLine 'useful for error checking

    Do
        OutputString = OutputString & Replace(Cells(RowNum, ColumnNum).Value, Chr(10), vbNewLine) & vbNewLine
        RowNum = RowNum + 1
    Loop Until IsEmpty(Cells(RowNum, ColumnNum))
        
    OutputString = OutputString & "Timeout 3"   'useful for errorchecking


    objFile.Write (OutputString)

    Set objFile = Nothing
    Set objFSO = Nothing

End Sub



'https://stackoverflow.com/questions/31359682/with-excel-vba-save-web-image-to-disk/31360105#31360105      , modified copypasta code, URLdownloadToFile

Sub dlStaplesImages()
    Dim rw As Long, lr As Long, ret As Long, sIMGDIR As String, sWAN As String, sLAN As String

    sIMGDIR = folderPath

    With ActiveSheet    '<-set this worksheet reference properly!
        lr = .Cells(Rows.Count, 1).End(xlUp).row
        For rw = 1 To lr 'rw to last row, assume first row is not header

            sWAN = .Cells(rw, 2).Value2
            sLAN = sIMGDIR & Chr(92) & Trim(Right(Replace(sWAN, Chr(47), Space(999)), 999))

            Debug.Print sWAN
            Debug.Print sLAN

            If CBool(Len(Dir(sLAN))) Then
                Call DeleteUrlCacheEntry(sLAN)
                Kill sLAN
            End If

            ret = URLDownloadToFile(0&, sWAN, sLAN, BINDF_GETNEWESTVERSION, 0&)
            
            'Imported code to output success / fail
            If ret = 0 Then
                Range("E" & rw).Value = "File successfully downloaded"
            Else
                Range("E" & rw).Value = "Unable to download the file"
            End If
            
            Next rw
    End With

End Sub

