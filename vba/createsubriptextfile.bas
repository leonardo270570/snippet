'to generate SRT file for subtitles
Public Sub GenerateSubripTextFile()
    Dim m_string, m_word As String
    Dim bolPunctuation As Boolean
    Dim m_value As Double
    Dim outputArr As Variant
    
    'write down all codes into text file
    Dim filePath As String
    Dim fso As FileSystemObject
    Dim fileStream As TextStream
    
    START_ROW_ON_SUBTITLE = 2
    SHEET_SUBTITLE_NAME = "convert"
    
    filePath = "D:\LeoFiles\Backup A50S\videography\code.srt"
    Set fso = New FileSystemObject
    ' Here the actual file is created and opened for write access
    Set fileStream = fso.CreateTextFile(filePath)
    'Set fileStream = fso.OpenTextFile(filePath, ForAppending, False, TristateMixed)

    m_workbook = ActiveWorkbook.Name
    m_row = START_ROW_ON_SUBTITLE 'starting row of the subtitle
    m_index = 1
    
    'looping from the module list on main page
    Do While Not IsEmpty(Sheets(SHEET_SUBTITLE_NAME).Cells(m_row, 1).Value)
        'm_index = Sheets(SHEET_SUBTITLE_NAME).Cells(m_row, 1).Value
        m_text = Sheets(SHEET_SUBTITLE_NAME).Cells(m_row, 13).Value ' the text into subriptext
        
        If m_text <> "" Then
            m_hour = 0 'Sheets(SHEET_SUBTITLE_NAME).Cells(m_row, 8).Value
            m_minute = 0 'Sheets(SHEET_SUBTITLE_NAME).Cells(m_row, 9).Value
            m_second = Sheets(SHEET_SUBTITLE_NAME).Cells(m_row - 1, 4).Value
            m_millisecond = Sheets(SHEET_SUBTITLE_NAME).Cells(m_row - 1, 5).Value
            
            m_secondto = Sheets(SHEET_SUBTITLE_NAME).Cells(m_row, 4).Value
            m_millisecondto = Sheets(SHEET_SUBTITLE_NAME).Cells(m_row, 5).Value
            
            fileStream.WriteLine m_index
            fileStream.WriteLine Mid(Trim(Str(100 + m_hour)), 2, 2) & ":" & _
                Mid(Trim(Str(100 + m_minute)), 2, 2) & ":" & _
                Mid(Trim(Str(100 + m_second)), 2, 2) & "," & _
                Mid(Trim(Str(1000 + m_millisecond)), 2, 3) & " --> " & _
                Mid(Trim(Str(100 + m_hour)), 2, 2) & ":" & _
                Mid(Trim(Str(100 + m_minute)), 2, 2) & ":" & _
                Mid(Trim(Str(100 + m_secondto)), 2, 2) & "," & _
                Mid(Trim(Str(1000 + m_millisecondto)), 2, 3)
            fileStream.WriteLine m_text
            fileStream.WriteLine
            
            m_index = m_index + 1
         End If
         
        m_row = m_row + 1
    Loop

    ' Close it, so it is not locked anymore
    fileStream.Close
End Sub
