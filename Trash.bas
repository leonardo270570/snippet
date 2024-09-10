Attribute VB_Name = "Trash"
'Created by Leonardo Sembiring, dated 13 June 2024
'Purpose : To calculate total commission

Public Sub CalculateCommission()
    Dim line As String
    Dim fso As Object
    Dim ts As Object
    Dim LineElements As Variant
    Dim Ind As Long
    Dim p As Long
    Dim Delimiter As String
    Dim FieldName() As String
    Dim m_string As String
    Dim m_array(18)
    
    Application.Calculation = xlCalculationManual
    
    strSheetMainVariable = "Main Variable"
    
    intRow = 7 'just change this row while you insert another row above
    strMainDirectory = Sheets(strSheetMainVariable).Cells(intRow, 2).Value 'main directory for reserving files
    strCurrentPeriod = Sheets(strSheetMainVariable).Cells(intRow + 1, 2).Value 'current period of valuation
    strPreviousPeriod = Sheets(strSheetMainVariable).Cells(intRow + 2, 2).Value 'previous period of valuation
    
    'read column of source
    For i = 1 To 18
        m_array(i) = Sheets(strSheetMainVariable).Cells(i + 6, 12).Value
    Next
    
    'open template
    strFile = "Reinsurance Credit Life Template.xlsx"
    'Workbooks.Open (strMainDirectory & "/reporting-template/" & strFile)

    'full name of variable_global.R
    Filename = strMainDirectory & "/" & Mid(strCurrentPeriod, 1, 6) & "/result/data-reinsurance-" & _
        Mid(strCurrentPeriod, 1, 6) & ".csv"
    
    Filename = "C:/SeaInsure/Actuarial-BAU/Reserve/202401/source/premium/Premium_Income-31-01-2024-01.csv"
    
    'file system object to manipulate text file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(Filename)
    Delimiter = "," 'delimiter follows R-generated file
    m_string = ""
    
    'name of the field
    line = ts.ReadLine
      
    FieldName = Split(line, Delimiter)
    For p = LBound(FieldName) To UBound(FieldName)
        Debug.Print p & " = " & FieldName(p)
    Next p
    
      
    intSheet = 1
    shtDetail = "Detail " & Trim(Str(intSheet))
    
    intSum = 0
    intRow = 2
    Do While ts.AtEndOfStream = False
        line = ts.ReadLine
        
        LineElements = Split(line, Delimiter)
'        For p = LBound(LineElements) To UBound(LineElemets)
'           Debug.Print FieldName(p) & " = " & p & " = " & LineElements(p)
'        Next p
'        If LineElements(7) = "IDGPPP2202" Or LineElements(7) = "IDGPSPP2302" Then
'            For i = 1 To 18
'                Workbooks(strFile).Sheets(shtDetail).Cells(intRow, i).Value = _
'                    LineElements(m_array(i))
'
'            Next
'            intRow = intRow + 1
'        End If
        If Val(LineElements(39)) <> 0 Then
            intSum = intSum + Val(LineElements(39)) + Val(LineElements(42))
        End If
        
        Ind = Ind + 1
      
'        If intRow = 500002 Then
'            intSheet = intSheet + 1
'            If intSheet > 3 Then Exit Do
'            shtDetail = "Detail " & Trim(Str(intSheet))
'            intRow = 2
'        End If
      
    Loop
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Debug.Print intSum
    
    Application.Calculation = xlCalculationAutomatic
End Sub


Sub ReadCSVFileSomeModelPoints()
    Dim line As String
    Dim fso As Object
    Dim ts As Object
    Dim LineElements As Variant
    Dim Ind As Long
    Dim p As Long
    Dim Delimiter As String
    Dim FieldName() As String
    Dim m_string As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Set ts = fso.OpenTextFile("C:\SeaInsure\Actuarial BAU\Reserve\202311\source\RI_Premium_QS_CLDT_2023-30-11-2023\RI_Premium_QS_CLDT_2023.csv")
    Set ts = fso.OpenTextFile("C:\SeaInsure\Actuarial BAU\Reserve\202311\source\GWP_and_Basic_Commission_UPR-30-11-2023\GWP_and_Basic_Commission_UPR-30-11-2023-00.csv")
    Delimiter = ","
    m_string = ""
    
    'title of the field
    line = ts.ReadLine
      
    FieldName = Split(line, Delimiter)
      
    Ind = 1
    Do While ts.AtEndOfStream = False
        line = ts.ReadLine
        
        If Ind = 10001 Then
            LineElements = Split(line, Delimiter)
            For p = LBound(LineElements) To UBound(LineElements)
               Debug.Print FieldName(p) & " = " & LineElements(p)
            Next p
            Exit Do
        End If
        
        Ind = Ind + 1
        'Debug.Print ""
    Loop
    'Debug.Print m_string
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

Sub ReadMergeFile()
    Dim line As String
    Dim fso As Object
    Dim ts As Object
    Dim LineElements As Variant
    Dim Ind As Long
    Dim p As Long
    Dim Delimiter As String
    Dim FieldName() As String
    Dim m_string As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Set ts = fso.OpenTextFile("C:\SeaInsure\Actuarial BAU\Reserve\202311\source\RI_Premium_QS_CLDT_2023-30-11-2023\RI_Premium_QS_CLDT_2023.csv")
    Set ts = fso.OpenTextFile("C:\SeaInsure\Actuarial-BAU\Reserve\202312\result\data-inforce-202312.csv")
    Delimiter = ","
    m_string = ""
    
    strSheetResult = "Result"
    
    '***** added by Leonardo Sembiring on Jan 16, 2024
    bolScreenUpdating = Application.ScreenUpdating
    bolStatusBar = Application.StatusBar
    bolCalculation = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.StatusBar = False
    Application.Calculation = xlCalculationManual


    'name of the field
    line = ts.ReadLine
      
    FieldName = Split(line, Delimiter)
    'For p = LBound(LineElements) To UBound(LineElements)
    '    FieldName(p) = LineElements(p)
    'Next p
      
      
    intRow = 1
    Do While ts.AtEndOfStream = False
        line = ts.ReadLine
        
        LineElements = Split(line, Delimiter)
'        For p = LBound(LineElements) To UBound(LineElements)
'           Debug.Print FieldName(p) & " = " & LineElements(p)
'        Next p
'
        'If InStr(1, m_string, Trim(LineElements(11)) & ";") = 0 Then
        '    m_string = m_string & Trim(LineElements(11)) & ";"
        'End If
'        If intRow = 1 Then
'            Debug.Print Mid(LineElements(10), 2, Len(LineElements(10)) - 1)
'            Exit Do
'        End If
        If LineElements(10) = "IDIPSLC2201" Then
            Sheets(strSheetResult).Cells(intRow, 1).Value = LineElements(3)
            Sheets(strSheetResult).Cells(intRow, 2).Value = LineElements(11)
            intRow = intRow + 1
            'Exit Do
        End If

'        If InStr(1, m_string, Trim(LineElements(14)) & ";") = 0 Then
'            m_string = m_string & Trim(LineElements(14)) & ";"
'        End If

      
      'Debug.Print ""
      
      'If Ind = 3 Then Exit Do
    Loop
    Debug.Print intRow 'm_string
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    'restore statusbar
    Application.StatusBar = bolStatusBar
    Application.ScreenUpdating = bolScreenUpdating
    Application.Calculation = bolCalculation
End Sub

'
'Option Explicit
'
'Sub Excel_vba_Shell_Command_Execute()
'    ' /C will execute the command and Terminate the window
'    MsgBox ExecShellCmd("C:\Users\leonardo.sembiring\AppData\Local\Programs\R\R-4.3.2\bin\RScript C:\SeaInsure\R\Leonardo\test.R")
'    '& "C:\Program Files\R\R-4.2.2\bin\Rscript.exe" "C:/SeaInsure/R/Leonardo/test.R"
'    'C:\Users\leonardo.sembiring\AppData\Local\Programs\R\R-4.3.2\bin\RScript C:\SeaInsure\R\Leonardo\test.R
'End Sub
'

Public Function ExecShellCmd(ByVal strRunExecution) As String
    Dim wsh As Object, wshOut As Object, sShellOut As String, sShellOutLine As String
    
    'strSheetMainVariable = "Main Variable"
    
    'strRunExecution = Sheets(strSheetMainVariable).Cells(10, 2).Value
    
    'Create object for Shell command execution
    Set wsh = CreateObject("WScript.Shell")

    'Run Excel VBA shell command and get the output string
    Set wshOut = wsh.exec(strRunExecution).stdout
    
    'Read each line of output from the Shell command & Append to Final Output Message
    While Not wshOut.AtEndOfStream
        sShellOutLine = wshOut.ReadLine
        If sShellOutLine <> "" Then
            sShellOut = sShellOut & sShellOutLine & vbCrLf
        End If
    Wend

    'Return the Output of Command Prompt
    'ExecShellCmd = "" 'sShellOut
    ExecShellCmd = sShellOut
End Function

Sub sampleProc()

    Dim command As String
    Dim wsh As Object
    Dim result As Integer
    
    'Set "command"
    strRun = "C:\Users\leonardo.sembiring\AppData\Local\Programs\R\R-4.3.2\bin\RScript C:\SeaInsure\Actuarial-BAU\Templates\R\main_program.R"

    
    Set wsh = CreateObject("WScript.Shell")
    
    'Run the "command"
    'Return
    '?0 : Succeeded
    '?1 : failed
    result = wsh.Run(strRun, WindowStyle:=0, WaitOnReturn:=True)  '"%ComSpec% /c " &
    
    If (result = 0) Then
        MsgBox ("Succeeded Run command.")
    Else
        MsgBox ("failed Run command.")
    End If
    
    Set wsh = Nothing
    
End Sub

Sub GetAllDirectories()
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    Dim i As Integer
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Get the folder object
    m_mainfolder = "H:\Project\r\prolife-actuarial\gpvreserve\assumption\"
    Set objFolder = objFSO.GetFolder(m_mainfolder)
    
    i = 1
    'loops through each file in the directory and prints their names and path
    For Each objSubFolder In objFolder.subfolders
    
        Set objChildFolder = objFSO.GetFolder(m_mainfolder & objSubFolder.Name & "\")
        
        For Each objFile In objChildFolder.Files
            'print folder name
            Debug.Print objFile.Name
            'print folder path
            Debug.Print objFile.Path
            i = i + 1
        Next objFile

        i = i + 1
    Next objSubFolder
    
    For Each objSubFolder In objFolder.Files
        'print folder name
        Debug.Print objSubFolder.Name
        'print folder path
        Debug.Print objSubFolder.Path
        i = i + 1
    Next objSubFolder

End Sub

Sub IndividualCells()
  Dim line As String
  Dim fso As Object
  Dim ts As Object
  Dim LineElements As Variant
  Dim Ind As Long
  Dim p As Long
  Dim Delimiter As String
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set ts = fso.OpenTextFile("H:\Project\r\prolife-actuarial\gpvreserve\assumption\mortality.csv")
  
  Delimiter = ","
  
  Ind = 1
  Do While ts.AtEndOfStream = False
    line = ts.ReadLine
    
    LineElements = Split(line, Delimiter)
    For p = LBound(LineElements) To UBound(LineElements)
       Cells(Ind, p + 1).Value = LineElements(p)
    Next p
    Ind = Ind + 1
  Loop
  
  ts.Close
  Set ts = Nothing
  Set fso = Nothing
End Sub

Sub WriteCSVFile()
    Dim fso As Object, ts As Object
     
    'Create the FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Create the TextStream
    Set ts = fso.CreateTextFile("H:\Project\r\prolife-actuarial\gpvreserve\hello.txt")
    'Write 2 lines ending with New Line character to text file
    ts.WriteLine "Hello World!"
    ts.WriteLine "Hello People!"
    ts.WriteLine "Hello People!"
    'Close the file
    ts.Close
     
    'Clean up memory
    Set fso = Nothing
    Set ts = Nothing
End Sub

