'helper for instruction
'strSheetName       : reference sheet where help listed
'intRow             : starting row of the list
'intCol             : starting column of the list
'intSpaceBetween    : number of columns between split of list
'intSplitCount      : number of split count

Public Sub ShowHelp(ByVal strSheetName As String, ByVal intRow As Integer, ByVal intCol As Integer, _
    ByVal intSpaceBetween As Integer, ByVal intSplitCount As Integer)
    strSheetItems = strSheetName
    
    Range(Cells(1, 3), Cells(intSplitCount, 200)).Select
    Selection.ClearContents
    Range("A1").Select

    intRowResult = 1
    intColResult = 3
    Do While Not IsEmpty(Sheets(strSheetItems).Cells(intRow, intCol).Value) 'reading index list
        intIndex = Sheets(strSheetItems).Cells(intRow, intCol).Value 'index
        strValue = Sheets(strSheetItems).Cells(intRow, intCol + 1).Value 'value
        
        ActiveSheet.Cells(intRowResult, intColResult).Value = intIndex
        ActiveSheet.Cells(intRowResult, intColResult + 1).Value = strValue
        
        intRow = intRow + 1
        intRowResult = intRowResult + 1
    
        If intRowResult > intSplitCount Then
            intRowResult = 1
            intColResult = intColResult + intSpaceBetween
        End If
    Loop
End Sub
