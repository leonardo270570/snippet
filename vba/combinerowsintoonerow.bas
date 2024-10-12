'combine some rows into one row : alt + enter
Public Sub WeeklyReport()
    Dim strParseWord(10) '10 could be extended
    
    intResultRow = 8
    intRow = 3
    Do While Not IsEmpty(ActiveSheet.Cells(intRow, 2).Value)
        strWord = ActiveSheet.Cells(intRow, 2).Value & " " & ActiveSheet.Cells(intRow, 3).Value
        strCheck = ActiveSheet.Cells(intRow, 8).Value
        strSolution = ActiveSheet.Cells(intRow, 4).Value
        
        If strCheck = "1" And strSolution <> "" Then

            intCountWord = 0 'counted number of words
            strOperator = "+" 'operation of numbers
            strOperatorEqual = "="
            
            'parsing the word into each part
            Do While True
                'to detect space in the word
                intInstr = InStr(1, strWord, " ")
                
                'parsing the words
                strParseWord(intCountWord) = Mid(strWord, 1, IIf(intInstr = 0, Len(strWord), intInstr - 1))
                
                If intInstr = 0 Then Exit Do
                
                strWord = Mid(strWord, intInstr + 1)
                intCountWord = intCountWord + 1
            Loop
            
            strWordWeek = ""
            For i = 0 To intCountWord
                strWordWeek = strWordWeek & strParseWord(i) & _
                    IIf(i = intCountWord, "", " " & IIf(i = intCountWord - 1, strOperatorEqual, strOperator) & " ")
            Next
            
            intInstr = InStr(1, strSolution, "=")
            strLetterSolution = Mid(strSolution, 1, intInstr - 2)
            strNumberSolution = Mid(strSolution, intInstr + 2)
            
            intLengthSolution = Len(strLetterSolution)
            strLetterWeek = ""
            intLength = Len(strWordWeek)
            For i = 1 To intLength
                strLetter = Mid(strWordWeek, i, 1)
                bolFind = False
                For j = 1 To intLengthSolution
                    If Mid(strLetterSolution, j, 1) = strLetter Then
                        strLetterWeek = strLetterWeek & Mid(strNumberSolution, j, 1)
                        bolFind = True
                        Exit For
                    End If
                Next
                If Not bolFind Then strLetterWeek = strLetterWeek & strLetter
            Next

            ActiveSheet.Cells(8, 10).Value = ActiveSheet.Cells(8, 10).Value & strWordWeek & _
                vbCrLf & strLetterWeek & vbCrLf & vbCrLf '& vbCrLf & strSolution
            
            
            intResultRow = intResultRow + 1
        End If
        
        intRow = intRow + 1
        
    Loop
    
End Sub '
