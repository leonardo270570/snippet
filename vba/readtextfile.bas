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
