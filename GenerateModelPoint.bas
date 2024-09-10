Attribute VB_Name = "GenerateModelPoint"
'Created by : Leonardo Sembiring
'Date       : 8 June 2023
'Source     : trad.xlsx, trad-k.xlsx, jiwa.xlsx
'Result     : csv files
  
Sub GenerateModelPoint()
    'sheet global variables
    m_wrk = ActiveWorkbook.Name
    m_mainsheet = "Main Variables"
    m_modelpointsheet = "Model Point"
    
    Sheets(m_modelpointsheet).Cells(12, 10).Value = "Running in progress…"

    m_valuationdate = Sheets(m_mainsheet).Cells(1, 2).Value
    m_excrateusd = Sheets(m_mainsheet).Cells(2, 2).Value
        
    'zerorized the starting count
    m_startrow = Sheets(m_modelpointsheet).Cells(15, 2).Value
    Do While Not IsEmpty(Sheets(m_modelpointsheet).Cells(m_startrow, 3).Value)
        Sheets(m_modelpointsheet).Cells(m_startrow, 5).Value = 0
        m_startrow = m_startrow + 1
    Loop
    
    'directory of model point files
    m_modelpointfolder = Sheets(m_modelpointsheet).Cells(1, 2).Value
    
    For j = 1 To 2
        m_source = Sheets(m_modelpointsheet).Cells(j + 5, 2).Value
        m_destination = Sheets(m_modelpointsheet).Cells(j + 5, 5).Value
        m_runstatus = Sheets(m_modelpointsheet).Cells(j + 5, 4).Value
        
        If m_runstatus = 1 Then
        
            m_count = Workbooks(m_source).Sheets.Count
                
            'run for all sheets in source file
            For i = 1 To m_count
                
                m_row = 2
                Do While True
                    m_policystatus = Workbooks(m_source).Sheets(i).Cells(m_row, 14).Value 'policy status
                    
                    If m_policystatus = "INFORCE" Then
                        m_policynumber = Workbooks(m_source).Sheets(i).Cells(m_row, 2).Value 'policynumber
                        m_productcode = Workbooks(m_source).Sheets(i).Cells(m_row, 15).Value 'product code
                        
                        m_entryage = Workbooks(m_source).Sheets(i).Cells(m_row, 11).Value 'age at entry
                        m_sex = Workbooks(m_source).Sheets(i).Cells(m_row, 10).Value 'sex
                        m_term = Workbooks(m_source).Sheets(i).Cells(m_row, 12).Value 'policy term
                        m_mode = Workbooks(m_source).Sheets(i).Cells(m_row, 16).Value  'payment mode
                        m_sa = Workbooks(m_source).Sheets(i).Cells(m_row, 17).Value  'sum assured
                        m_premium = Workbooks(m_source).Sheets(i).Cells(m_row, 19).Value  'annual premium
                        m_commencedate = Workbooks(m_source).Sheets(i).Cells(m_row, 8).Value  'commencement date
                        m_paymentterm = Workbooks(m_source).Sheets(i).Cells(m_row, 13).Value  'payment term
                        'm_crbenefit = 0 'initialize critical illness benefit
                        'm_loanpc = 0 'initialize the loan pc
                        m_diff = (Year(m_valuationdate) - Year(m_commencedate)) * 12 + (Month(m_valuationdate) - Month(m_commencedate))
                        
                        'payment mode
                        Sheets(m_modelpointsheet).Cells(17, 2).Value = m_mode
                        m_paymentmode = Sheets(m_modelpointsheet).Cells(18, 2).Value
                        m_mutiplier = Sheets(m_modelpointsheet).Cells(19, 2).Value
                        
                        'fill in the product code to get total current data
                        Sheets(m_modelpointsheet).Cells(11, 2).Value = m_productcode
                        m_listrow = Sheets(m_modelpointsheet).Cells(14, 2).Value
                        m_destinationsheet = Sheets(m_modelpointsheet).Cells(12, 2).Value
                        
                        'to add count for count policies
                        Sheets(m_modelpointsheet).Cells(m_listrow, 5).Value = _
                            Sheets(m_modelpointsheet).Cells(m_listrow, 5).Value + 1
                        
                        'need a row position to write into product sheet
                        m_destinationrow = Sheets(m_modelpointsheet).Cells(m_listrow, 5).Value + 1
            
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 1).Value = "'" & m_policynumber
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 2).Value = m_productcode
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 3).Value = m_entryage
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 4).Value = m_sex
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 5).Value = m_term
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 6).Value = m_diff + 1
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 7).Value = Year(m_commencedate)
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 8).Value = Month(m_commencedate)
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 9).Value = m_paymentterm
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 10).Value = m_premium * m_mutiplier
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 11).Value = m_paymentmode
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 12).Value = m_sa
                    End If
                    
                    If IsEmpty(Workbooks(m_source).Sheets(i).Cells(m_row + 1, 1).Value) Then
                        Exit Do
                    End If
        
                    m_row = m_row + 1
                Loop
            Next
        End If
    Next
    
    'for JIWA
    For j = 3 To 3
        m_source = Sheets(m_modelpointsheet).Cells(j + 5, 2).Value
        m_destination = Sheets(m_modelpointsheet).Cells(j + 5, 5).Value
        m_runstatus = Sheets(m_modelpointsheet).Cells(j + 5, 4).Value
        
        If m_runstatus = 1 Then
            m_count = Workbooks(m_source).Sheets.Count
                
            'run for all sheets in source file
            For i = 1 To m_count
                
                m_row = 2
                Do While True
                    m_policystatus = Workbooks(m_source).Sheets(i).Cells(m_row, 12).Value 'policy status
                    
                    If m_policystatus = "INFORCE" Then
                        m_policynumber = Workbooks(m_source).Sheets(i).Cells(m_row, 2).Value 'policynumber
                        m_productcode = "JIWA" 'Workbooks(m_source).Sheets(i).Cells(m_row, 15).Value 'product code
                        
                        m_entryage = Workbooks(m_source).Sheets(i).Cells(m_row, 9).Value 'age at entry
                        m_sex = Workbooks(m_source).Sheets(i).Cells(m_row, 8).Value 'sex
                        m_term = 10 'Workbooks(m_source).Sheets(i).Cells(m_row, 12).Value 'policy term
                        m_mode = "TAHUNAN" 'Workbooks(m_source).Sheets(i).Cells(m_row, 16).Value  'payment mode
                        m_sa = Workbooks(m_source).Sheets(i).Cells(m_row, 14).Value  'sum assured
                        m_premium = Workbooks(m_source).Sheets(i).Cells(m_row, 15).Value  'annual premium
                        m_commencedate = Workbooks(m_source).Sheets(i).Cells(m_row, 10).Value  'commencement date
                        m_paymentterm = 5 'Workbooks(m_source).Sheets(i).Cells(m_row, 13).Value  'payment term
                        'm_crbenefit = 0 'initialize critical illness benefit
                        'm_loanpc = 0 'initialize the loan pc
                        m_diff = (Year(m_valuationdate) - Year(m_commencedate)) * 12 + (Month(m_valuationdate) - Month(m_commencedate))
                        
                        'payment mode
                        Sheets(m_modelpointsheet).Cells(17, 2).Value = m_mode
                        m_paymentmode = Sheets(m_modelpointsheet).Cells(18, 2).Value
                        
                        'fill in the product code to get total current data
                        Sheets(m_modelpointsheet).Cells(11, 2).Value = m_productcode
                        m_listrow = Sheets(m_modelpointsheet).Cells(14, 2).Value
                        m_destinationsheet = Sheets(m_modelpointsheet).Cells(12, 2).Value
                        
                        'to add count for count policies
                        Sheets(m_modelpointsheet).Cells(m_listrow, 5).Value = _
                            Sheets(m_modelpointsheet).Cells(m_listrow, 5).Value + 1
                        
                        'need a row position to write into product sheet
                        m_destinationrow = Sheets(m_modelpointsheet).Cells(m_listrow, 5).Value + 1
            
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 1).Value = "'" & m_policynumber
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 2).Value = m_productcode
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 3).Value = m_entryage
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 4).Value = m_sex
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 5).Value = m_term
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 6).Value = m_diff + 1
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 7).Value = Year(m_commencedate)
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 8).Value = Month(m_commencedate)
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 9).Value = m_paymentterm
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 10).Value = m_premium
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 11).Value = m_paymentmode
                        Workbooks(m_destination).Worksheets(m_destinationsheet). _
                            Cells(m_destinationrow, 12).Value = m_sa
                    End If
                    
                    If IsEmpty(Workbooks(m_source).Sheets(i).Cells(m_row + 1, 1).Value) Then
                        Exit Do
                    End If
        
                    m_row = m_row + 1
                Loop
            Next
        End If
    Next

'    Workbooks(m_destination).Activate
'    m_count = ActiveWorkbook.Sheets.Count
'
'    'to generate the model point files
'    For j = 1 To m_count
'        Sheets(j).Select
'        Sheets(j).Cells(1, 2).Value = m_valuationdate
'        Sheets(j).Cells(2, 2).Value = m_excrateusd
'        'Sheets(j).Cells(3, 2).Value = m_excratesgd
'        m_1 = Sheets(j).Name
'        m_row = Sheets(j).Cells(6, 2).Value
'        m_col = Sheets(j).Cells(7, 2).Value
'        m_filename = Sheets(j).Cells(4, 2).Value
'        m_ext = Sheets(j).Cells(5, 2).Value
'
'        Filename = m_dir & m_filename & "." & m_ext
'
'        Set fs = CreateObject("Scripting.FileSystemObject")
'        Set a = fs.CreateTextFile(Filename, True)
'
'        For q = 1 To m_row
'            m_writeline = ""
'            For ai = 1 To m_col
'                m_writeline = m_writeline & Cells(q + 10, ai) & ","
'            Next
'            a.WriteLine Mid(m_writeline, 1, Len(m_writeline) - 1)
'        Next q
'
'        a.Close
'    Next

    Sheets(m_modelpointsheet).Cells(12, 10).Value = ""
End Sub


