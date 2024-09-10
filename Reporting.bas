Attribute VB_Name = "Reporting"
'Created by Leonardo Sembiring, dated 27 February 2024
'Purpose : To generate reinsurance report

Public Sub CreditLifeReinsuranceReport()
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
    Workbooks.Open (strMainDirectory & "/reporting-template/" & strFile)

    'full name of variable_global.R
    Filename = strMainDirectory & "/" & Mid(strCurrentPeriod, 1, 6) & "/result/data-reinsurance-" & _
        Mid(strCurrentPeriod, 1, 6) & ".csv"

    'file system object to manipulate text file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(Filename)
    Delimiter = ";" 'delimiter follows R-generated file
    m_string = ""
    
    'name of the field
    line = ts.ReadLine
      
    FieldName = Split(line, Delimiter)
    For p = LBound(FieldName) To UBound(FieldName)
        Debug.Print p & " = " & FieldName(p)
    Next p
    
      
    intSheet = 1
    shtDetail = "Detail " & Trim(Str(intSheet))
    
    intRow = 2
    Do While ts.AtEndOfStream = False
        line = ts.ReadLine
        
        LineElements = Split(line, Delimiter)
'        For p = LBound(LineElements) To UBound(LineElements)
'           Debug.Print FieldName(p) & " = " & p & " = " & LineElements(p)
'        Next p
        If LineElements(7) = "IDGPPP2202" Or LineElements(7) = "IDGPSPP2302" Then
            For i = 1 To 18
                Workbooks(strFile).Sheets(shtDetail).Cells(intRow, i).Value = _
                    LineElements(m_array(i))
                
            Next
            intRow = intRow + 1
        End If

        Ind = Ind + 1
      
        If intRow = 500002 Then
            intSheet = intSheet + 1
            If intSheet > 3 Then Exit Do
            shtDetail = "Detail " & Trim(Str(intSheet))
            intRow = 2
        End If
      
    Loop
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Application.Calculation = xlCalculationAutomatic
End Sub


Public Sub TermLifeReinsuranceReport()
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
        m_array(i) = Sheets(strSheetMainVariable).Cells(i + 6, 14).Value
    Next
    
    'open template
    strFile = "Reinsurance Term Life Template.xlsx"
    Workbooks.Open (strMainDirectory & "/reporting-template/" & strFile)

    'full name of variable_global.R
    Filename = strMainDirectory & "/" & Mid(strCurrentPeriod, 1, 6) & "/result/data-reinsurance-" & _
        Mid(strCurrentPeriod, 1, 6) & ".csv"

    'file system object to manipulate text file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(Filename)
    Delimiter = ";" 'delimiter follows R-generated file
    m_string = ""
    
    'name of the field
    line = ts.ReadLine
    'MsgBox line
    
    FieldName = Split(line, Delimiter)
    For p = LBound(FieldName) To UBound(FieldName)
        Debug.Print p & " = " & FieldName(p)
    Next p
    
      
    intSheet = 1
    shtDetail = "Detail " & Trim(Str(intSheet))
    
    intRow = 2
    Do While ts.AtEndOfStream = False
        line = ts.ReadLine
        'MsgBox line
        
        LineElements = Split(line, Delimiter)
'        For p = LBound(LineElements) To UBound(LineElements)
'           Debug.Print FieldName(p) & " = " & p & " = " & LineElements(p)
'        Next p
        If LineElements(7) = "IDIPSLC2201" Then
            For i = 1 To 18
                Workbooks(strFile).Sheets(shtDetail).Cells(intRow, i).Value = _
                    LineElements(m_array(i))
                
            Next
            intRow = intRow + 1
        End If

        Ind = Ind + 1
      
        If intRow = 500002 Then
            intSheet = intSheet + 1
            If intSheet > 3 Then Exit Do
            shtDetail = "Detail " & Trim(Str(intSheet))
            intRow = 2
        End If
      
    Loop
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Application.Calculation = xlCalculationAutomatic
End Sub


Public Sub ThreePAReinsuranceReport()
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
        m_array(i) = Sheets(strSheetMainVariable).Cells(i + 6, 14).Value
    Next
    
    'open template
    strFile = "Reinsurance 3PA Template.xlsx"
    Workbooks.Open (strMainDirectory & "/reporting-template/" & strFile)

    'full name of variable_global.R
    Filename = strMainDirectory & "/" & Mid(strCurrentPeriod, 1, 6) & "/result/data-reinsurance-" & _
        Mid(strCurrentPeriod, 1, 6) & ".csv"

    'file system object to manipulate text file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(Filename)
    Delimiter = ";" 'delimiter follows R-generated file
    m_string = ""
    
    'name of the field
    line = ts.ReadLine
      
    FieldName = Split(line, Delimiter)
    For p = LBound(FieldName) To UBound(FieldName)
        Debug.Print p & " = " & FieldName(p)
    Next p
    
      
    intSheet = 1
    shtDetail = "Detail " & Trim(Str(intSheet))
    
    intRow = 2
    Do While ts.AtEndOfStream = False
        line = ts.ReadLine
        
        LineElements = Split(line, Delimiter)
'        For p = LBound(LineElements) To UBound(LineElements)
'           Debug.Print FieldName(p) & " = " & p & " = " & LineElements(p)
'        Next p
        If LineElements(7) = "IDHISMCP2201" Or LineElements(7) = "IDHISMTD2201" Or LineElements(7) = "IDPASPA2201" Then
            For i = 1 To 18
                Workbooks(strFile).Sheets(shtDetail).Cells(intRow, i).Value = _
                    LineElements(m_array(i))
                
            Next
            intRow = intRow + 1
        End If

        Ind = Ind + 1
      
        If intRow = 500002 Then
            intSheet = intSheet + 1
            If intSheet > 3 Then Exit Do
            shtDetail = "Detail " & Trim(Str(intSheet))
            intRow = 2
        End If
      
    Loop
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Application.Calculation = xlCalculationAutomatic
End Sub


Public Sub CriticalIllnessReinsuranceReport()
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
        m_array(i) = Sheets(strSheetMainVariable).Cells(i + 6, 16).Value
    Next
    
    'open template
    strFile = "Reinsurance Critical Illness Template.xlsx"
    Workbooks.Open (strMainDirectory & "/reporting-template/" & strFile)

    'full name of variable_global.R
    Filename = strMainDirectory & "/" & Mid(strCurrentPeriod, 1, 6) & "/result/data-reinsurance-" & _
        Mid(strCurrentPeriod, 1, 6) & ".csv"

    'file system object to manipulate text file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(Filename)
    Delimiter = ";" 'delimiter follows R-generated file
    m_string = ""
    
    'name of the field
    line = ts.ReadLine
      
    FieldName = Split(line, Delimiter)
    For p = LBound(FieldName) To UBound(FieldName)
        Debug.Print p & " = " & FieldName(p)
    Next p
    
      
    intSheet = 1
    shtDetail = "Detail " & Trim(Str(intSheet))
    
    intRow = 2
    Do While ts.AtEndOfStream = False
        line = ts.ReadLine
        
        LineElements = Split(line, Delimiter)
'        For p = LBound(LineElements) To UBound(LineElements)
'           Debug.Print FieldName(p) & " = " & p & " = " & LineElements(p)
'        Next p
        If LineElements(7) = "IDIPSMCI2201" Then 'IDIPSMCI2201
            For i = 1 To 18
                Workbooks(strFile).Sheets(shtDetail).Cells(intRow, i).Value = _
                    LineElements(m_array(i))
                
            Next
            intRow = intRow + 1
        End If

        Ind = Ind + 1
      
        If intRow = 500002 Then
            intSheet = intSheet + 1
            If intSheet > 3 Then Exit Do
            shtDetail = "Detail " & Trim(Str(intSheet))
            intRow = 2
        End If
      
    Loop
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Application.Calculation = xlCalculationAutomatic
End Sub

