Sub GenerateFakeData(ByVal m_modulename As String, ByVal m_row As Integer)
    'Dim arr() As String
    'write down all codes into text file
    Dim filePath As String
    Dim fso As FileSystemObject
    Dim fileStream As TextStream
    
    m_space = " "
    'main page level info
    m_sheetname = SHEET_TABLE_NAME
    m_modulerow = Sheets(SHEET_MAIN_PAGE_NAME).Cells(m_row, 3).Value 'table row reference for module
    m_tablename = Sheets(m_sheetname).Cells(m_modulerow, 1).Value 'table name from sheet Table
    
    'location of the file : E:\Project\insurance\fakeserver\db.json
    filePath = FAKE_DATA_FOLDER & "db.json"
    Set fso = New FileSystemObject
    ' Here the actual file is created and opened for write access
    Set fileStream = fso.CreateTextFile(filePath)
    
    'evaluates total field in json file, needed to remove the comma sign for the last field
    m_totalfield = 0
    m_row = m_modulerow
    Do While Not IsEmpty(Sheets(m_sheetname).Cells(m_row, FAKE_DATA_CHECK_COLUMN).Value)
        m_checked = Sheets(m_sheetname).Cells(m_row, FAKE_DATA_CHECK_COLUMN).Value 'checked or not
        If m_checked = 1 Then
            m_totalfield = m_totalfield + 1
        End If
        m_row = m_row + 1
    Loop

    'opening curl
    fileStream.WriteLine "{"
    fileStream.WriteLine INDENT_1 & """" & LCase(m_modulename) & """" & ": ["

    m_col = 13 'start column of field fake data
    m_datacount = 3 'number of dummy data
    
    For m_count = 1 To m_datacount
    
        m_row = m_modulerow
        m_fieldcount = 1
        fileStream.WriteLine INDENT_2 & "{"
        
        Do While Not IsEmpty(Sheets(m_sheetname).Cells(m_row, FAKE_DATA_CHECK_COLUMN).Value)
            
            m_variablename = Sheets(m_sheetname).Cells(m_row, 6).Value 'field name of the associated table
            m_datatype = Sheets(m_sheetname).Cells(m_row, 7).Value 'type of the field
            m_value = Sheets(m_sheetname).Cells(m_row, m_col).Value 'value of the field
            m_checked = Sheets(m_sheetname).Cells(m_row, FAKE_DATA_CHECK_COLUMN).Value 'checked or not
            
            If m_checked = 1 Then
                If m_fieldcount = m_totalfield Then m_addcomma = "" Else m_addcomma = ","
                
                Call WriteJsonField(fileStream, 2, m_variablename, ValueByDataType(m_datatype, m_value), m_addcomma)
                m_fieldcount = m_fieldcount + 1
            End If
            
            m_row = m_row + 1
        Loop
        
        'remove comma sign for the last row data
        If m_count = m_datacount Then
            fileStream.WriteLine INDENT_2 & "}"
        Else
            fileStream.WriteLine INDENT_2 & "},"
        End If
        m_col = m_col + 1
    Next
    
    'closing curl
    fileStream.WriteLine INDENT_1 & "]"
    fileStream.WriteLine "}"
    
    ' Close it, so it is not locked anymore
    fileStream.Close

End Sub

Public Function ValueByDataType(ByVal m_string, ByVal m_value) As String
    
    m_returnstring = ""
    
    If InStr(1, m_string, "nvarchar") Then m_string = "nvarchar"
    If InStr(1, m_string, "datetime") Then m_string = "datetime"
    
    Select Case m_string
        Case "int"
            m_returnstring = "number"
        Case "nvarchar"
            m_returnstring = "string"
        Case "bit"
            m_returnstring = "boolean"
        Case "uniqueidentifier"
            m_returnstring = "number"
        Case "datetime"
            m_returnstring = "number"
        Case Else
            m_returnstring = "number"
    End Select
    
    m_returnstring = IIf(m_returnstring = "string", """" & m_value & """", m_value)
    
    ValueByDataType = m_returnstring
End Function
