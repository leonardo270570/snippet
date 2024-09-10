Attribute VB_Name = "PublicFunction"

'Convert CSV file from Excel into R format (UTF-8)
Sub convertTxttoUTF(sInFilePath As String, sOutFilePath As String)
    Dim objFS  As Object
    Dim iFile       As Double
    Dim sFileData   As String
    
    'Init
    iFile = FreeFile
    Open sInFilePath For Input As #iFile
        sFileData = Input$(LOF(iFile), iFile)
        sFileData = sFileData & vbCrLf
    Close iFile
    
    'Open & Write
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText sFileData
    
    'Save & Close
    objFS.SaveToFile sOutFilePath, 2   '2: Create Or Update
    objFS.Close
    
    'Completed
    'Application.StatusBar = "Completed"
End Sub

'to delete file log
Public Sub DeleteFileLog()
    On Error Resume Next
    
    Dim strFileName As String
    
    strSheetRunLog = "Run Log"
    strFileName = Sheets(strSheetRunLog).Cells(1, 2).Value
    
    Kill strFileName
End Sub

'to read file log and showing in the spreadheet
Public Sub ReadLogFile()
    On Error Resume Next

    Dim line As String
    Dim fso As Object
    Dim ts As Object
    Dim LineElements As Variant
    Dim Delimiter As String
    Dim inRow As Integer
    
    Dim strFileName As String
    
    strSheetRunLog = "Run Log"
    strFileName = Sheets(strSheetRunLog).Cells(1, 2).Value
    strFileName = Replace(strFileName, "/", "\")

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'clear previous comments
    Range("A8:B500").Select
    Selection.ClearContents
    Range("A8").Select

    'check the file log exist or not
    FileExists = fso.FileExists(strFileName)
    If Not FileExists Then
        Set fso = Nothing
        Exit Sub
    End If
    
    Set ts = fso.OpenTextFile(strFileName)
    Delimiter = ";"
    line = ts.ReadLine 'skip the title of the table

    intRow = 8 'starting row of the log to be shown
    Do While ts.AtEndOfStream = False
        line = ts.ReadLine

        LineElements = Split(line, Delimiter)
        Sheets(strSheetRunLog).Cells(intRow, 1).Value = LineElements(0)
        Sheets(strSheetRunLog).Cells(intRow, 2).Value = LineElements(1)

        intRow = intRow + 1
    Loop

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

'parsing text from double quotes string
Public Function ParseText(ByVal strText As String)
    ParseText = Mid(strText, 2, Len(strText) - 2)
End Function
