'read value of index
Public Function ValueOfIndex(ByVal intIndexSearch As Integer, ByVal strSheetName As String, _
    ByVal intRow As Integer, ByVal intCol As Integer)
    
    strValueSearch = ""
    Do While Not IsEmpty(Sheets(strSheetName).Cells(intRow, intCol).Value) 'reading index list
        intIndex = Sheets(strSheetName).Cells(intRow, intCol).Value 'index
        strValue = Sheets(strSheetName).Cells(intRow, intCol + 1).Value 'value
        If intIndexSearch = intIndex Then
            strValueSearch = strValue
            Exit Do
        End If
        intRow = intRow + 1
    Loop
    
    ValueOfIndex = strValueSearch
End Function
