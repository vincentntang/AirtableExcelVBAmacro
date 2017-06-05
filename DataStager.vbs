
Option Explicit

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

    'Global Variables for passing values b/w subs
    Dim myPath As String
    Dim folderPath As String
    Dim folderLocation As Variant





Sub airtableCleaner()
    Dim argCounter As Integer
    Dim Answer As VbMsgBoxResult

    folderPath = Application.ActiveWorkbook.Path 'Example C:/downloads
    myPath = Application.ActiveWorkbook.FullName 'Example C:/downloads/book1.csv

    'Ask user if they want to run macro
    Answer = MsgBox("Run? Airtable - 1: primaryKey, 2: one image attachment)", vbYesNo, "Run Macro")
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
    
    '-------------------------------------------------------------------------
    'Cleanup Broken links %5B1%5D in Column C  (only needed for extremePictureFinder, not with ExcelVBA URLimagefinder
    '-------------------------------------------------------------------------
    'Columns("C:C").Select
    'Range("C40").Activate
    'Selection.Replace What:="%5B1%5D", Replacement:="[1]", LookAt:=xlPart, _
    'SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    'ReplaceFormat:=False
    '-------------------------------------------------------------------------

    
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
    
    'Image downloader to source folder
    Call dlStaplesImages
    
    'Make the batch files using row data col D
    Call ExportRangetoBatch
    
    'Ask user to run bat file now or later
    Shell "cmd.exe /k cd " & folderPath & " && newcurl.bat"

    
    End If
End Sub

'https://superuser.com/questions/1045707/create-bat-file-with-excel-data-with-vba    , modified copypasta code

Sub ExportRangetoBatch()

    Dim ColumnNum: ColumnNum = 4   ' Column D
    Dim RowNum: RowNum = 1          ' Row to start on
    Dim objFSO, objFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(folderPath & "\newcurl.bat")    'Output Path

    Dim OutputString: OutputString = ""

    Do
        OutputString = OutputString & Replace(Cells(RowNum, ColumnNum).Value, Chr(10), vbNewLine) & vbNewLine
        RowNum = RowNum + 1
    Loop Until IsEmpty(Cells(RowNum, ColumnNum))

    objFile.Write (OutputString)

    Set objFile = Nothing
    Set objFSO = Nothing

End Sub



'https://stackoverflow.com/questions/31359682/with-excel-vba-save-web-image-to-disk/31360105#31360105      , modified copypasta code

Sub dlStaplesImages()
    Dim rw As Long, lr As Long, ret As Long, sIMGDIR As String, sWAN As String, sLAN As String

    sIMGDIR = folderPath
    'If Dir(sIMGDIR, vbDirectory) = "" Then MkDir sIMGDIR

    With ActiveSheet    '<-set this worksheet reference properly!
        lr = .Cells(Rows.Count, 1).End(xlUp).Row
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
            
            '.Cells(rw, 5) = ret
            Next rw
    End With

End Sub



