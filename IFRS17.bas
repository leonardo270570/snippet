Attribute VB_Name = "IFRS17"
Public Sub PrepareUPRGOC()
    Dim wb As Workbook
    Dim strFolder As String
    Dim strFile As String
    
    strWorkbook = ActiveWorkbook.Name
    strSheetResult = "Result"

    strFolder = "C:\Actuarial-BAU\IFRS17\"
    strFile = Dir(strFolder & "*.xls*")
        
    Do While strFile <> ""
        
        If Mid(strFile, 1, 24) = "Portfolio Inforce_Group_" Then
        'If Mid(strFile, 1, 7) = "Claims_" Then
            Set wb = Workbooks.Open(strFolder & strFile)
            
            '1222 0123.... 0124
            strGroup = Mid(strFile, 25, 4)
            'column of writing result
            intColResult = Int(Mid(strGroup, 1, 2)) + 12 * (Int(Mid(strGroup, 3, 2)) - 22) - 9
            shtData = "Data IF"
            
            intProductCode = 70: intIssueYear = 70: intUpr = 70: intRiUpr = 70
            For i = 1 To 70
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Issue Year" Then intIssueYear = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "UPR" Then intUpr = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "RI UPR" Then intRiUpr = i
            Next
            Debug.Print intProductCode
            Debug.Print intIssueYear
            Debug.Print intUpr
            Debug.Print intRiUpr
            
            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value
                intReadIssueYear = Val(Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value)
                dblReadUpr = Workbooks(strFile).Sheets(shtData).Cells(intRow, intUpr).Value
                dblReadRiUpr = Workbooks(strFile).Sheets(shtData).Cells(intRow, intRiUpr).Value
                
                For i = 1 To 8
                    If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                        'row of writing result
                        intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                        Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblReadRiUpr 'dblReadUpr
                    End If
                Next
                intRow = intRow + 1
            Loop
            
            If sheetExists Then
                shtDataSCL = "SCL DS"
                intProductCode = 70: intIssueYear = 70: intUpr = 70: intRiUpr = 70
                For i = 1 To 70
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "product_code" Then intProductCode = i
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "issue_date" Then intIssueYear = i
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "UPR" Then intUpr = i
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "RI UPR" Then intRiUpr = i
                Next
                Debug.Print intProductCode
                Debug.Print intIssueYear
                Debug.Print intUpr
                Debug.Print intRiUpr
                
                intRow = 4
                Do While Not IsEmpty(Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intProductCode).Value)
                    strReadProductCode = Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intProductCode).Value
                    intReadIssueYear = Year(Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intIssueYear).Value)
                    dblReadUpr = Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intUpr).Value
                    dblReadRiUpr = Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intRiUpr).Value
                    
                    For i = 1 To 8
                        If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                            'row of writing result
                            intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                                Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblReadRiUpr 'dblReadUpr
                        End If
                    Next
                    intRow = intRow + 1
                Loop
            End If

            Workbooks(strFile).Close SaveChanges:=False
        End If
        
        If Mid(strFile, 1, 29) = "Portfolio Inforce_Individual_" Then
            Set wb = Workbooks.Open(strFolder & strFile)
            
            '1222 0123.... 0124
            strIndividual = Mid(strFile, 30, 4)
            'column of writing result
            intColResult = Int(Mid(strIndividual, 1, 2)) + 12 * (Int(Mid(strIndividual, 3, 2)) - 22) - 9
            
            shtData = "Data IF"
            intProductCode = 70: intIssueYear = 70: intUpr = 70: intRiUpr = 70
            For i = 1 To 70
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Issued Year" Then intIssueYear = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "UPR" Then intUpr = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "RI UPR" Then intRiUpr = i
            Next
            Debug.Print intProductCode
            Debug.Print intIssueYear
            Debug.Print intUpr
            Debug.Print intRiUpr
            
            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value
                intReadIssueYear = Val(Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value)
                dblReadUpr = Workbooks(strFile).Sheets(shtData).Cells(intRow, intUpr).Value
                dblReadRiUpr = Workbooks(strFile).Sheets(shtData).Cells(intRow, intRiUpr).Value
                
                For i = 1 To 8
                    If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                        'row of writing result
                        intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                        Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblReadRiUpr 'dblReadUpr
                    End If
                Next
                intRow = intRow + 1
            Loop
            
            Workbooks(strFile).Close SaveChanges:=False
        End If

        strFile = Dir
    Loop
End Sub


Public Sub PrepareOCGOC()
    Dim wb As Workbook
    Dim strFolder As String
    Dim strFile As String
    
    strWorkbook = ActiveWorkbook.Name
    strSheetResult = "Result"

    strFolder = "C:\Actuarial-BAU\IFRS17\"
    strFile = Dir(strFolder & "*.xls*")
        
    Do While strFile <> ""
        
        If Mid(strFile, 1, 7) = "Claims_" Then
            Set wb = Workbooks.Open(strFolder & strFile)
                        
            '1222 0123.... 0124
            strClaim = Mid(strFile, 8, 4)
            'column of writing result
            intColResult = Int(Mid(strClaim, 1, 2)) + 12 * (Int(Mid(strClaim, 3, 2)) - 22) - 9
            shtData = "Claims"
            
            intProductCode = 70: intIssueYear = 70: intClaimStatus = 70: intOSClaim = 70: intRiOSClaim = 70
            For i = 1 To 70
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Policy Effective Date" Then intIssueYear = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Claim Status" Then intClaimStatus = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Claim Outstanding Reserve" Then intOSClaim = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Claim RI Outstanding Recovery" Then intRiOSClaim = i
            Next
            Debug.Print intProductCode
            Debug.Print intIssueYear
            Debug.Print intOSClaim
            Debug.Print intRiOSClaim
            
            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value
                If Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value = "" Then
                    intReadIssueYear = 2022
                Else
                    intReadIssueYear = Year(Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value)
                End If
                strClaimStatus = Workbooks(strFile).Sheets(shtData).Cells(intRow, intClaimStatus).Value
                dblReadOSClaim = Workbooks(strFile).Sheets(shtData).Cells(intRow, intOSClaim).Value
                dblReadRiOSClaim = Workbooks(strFile).Sheets(shtData).Cells(intRow, intRiOSClaim).Value
                
                If strClaimStatus = "Pending" Then
                    For i = 1 To 8
                        If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                            'row of writing result
                            intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                                Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblReadRiOSClaim 'dblReadOSClaim ' ' '
                        End If
                    Next
                End If
                
                intRow = intRow + 1
            Loop
            

            Workbooks(strFile).Close SaveChanges:=False
        End If
        
        strFile = Dir
    Loop
End Sub


Public Sub PrepareDACGOC()
    Dim wb As Workbook
    Dim strFolder As String
    Dim strFile As String
    
    strWorkbook = ActiveWorkbook.Name
    strSheetResult = "Result"

    strFolder = "C:\Actuarial-BAU\IFRS17\"
    strFile = Dir(strFolder & "*.xls*")
        
    Do While strFile <> ""
        
        If Mid(strFile, 1, 24) = "Portfolio Inforce_Group_" Then
        'If Mid(strFile, 1, 7) = "Claims_" Then
            Set wb = Workbooks.Open(strFolder & strFile)

            'to check if SLC DS sheet exists
            sheetExists = False
            For Each ws In Workbooks(strFile).Worksheets
                If ws.Name = "SCL DS" Then
                    sheetExists = True
                    'MsgBox strFile
                    Exit For
                End If
            Next ws

            '1222 0123.... 0124
            strGroup = Mid(strFile, 25, 4)
            'column of writing result
            intColResult = Int(Mid(strGroup, 1, 2)) + 12 * (Int(Mid(strGroup, 3, 2)) - 22) - 9
            shtData = "Data IF"
            
            intProductCode = 70: intIssueYear = 70: intPD = 70: intCommission = 70: intEarnedPremium = 70: intPremium = 70
            For i = 1 To 70
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Issue Year" Then intIssueYear = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Premium Discount" Then intPD = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Commission" Then intCommission = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Earned Premium" Then intEarnedPremium = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Premium" Then intPremium = i
            Next
            
            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value
                intReadIssueYear = Val(Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value)
                dblPremiumDiscount = Workbooks(strFile).Sheets(shtData).Cells(intRow, intPD).Value
                dblCommission = Workbooks(strFile).Sheets(shtData).Cells(intRow, intCommission).Value
                dblEarnedPremium = Workbooks(strFile).Sheets(shtData).Cells(intRow, intEarnedPremium).Value
                dblPremium = Workbooks(strFile).Sheets(shtData).Cells(intRow, intPremium).Value
                
                For i = 1 To 8
                    If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                        'row of writing result
                        intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                        'times zero for premium discount
                        dblDac = (dblPremiumDiscount * 0 + dblCommission) * (1 - dblEarnedPremium / dblPremium)
                        Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblDac
                    End If
                Next
                intRow = intRow + 1
            Loop
            
            If sheetExists Then
                shtDataSCL = "SCL DS"
                intProductCode = 70: intIssueYear = 70: intPremiumDiscount = 70: intRemainingPOI = 70
                For i = 1 To 70
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "product_code" Then intProductCode = i
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "issue_date" Then intIssueYear = i
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "premium_discount" Then intPremiumDiscount = i
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "remaining_poi" Then intRemainingPOI = i
                Next
                Debug.Print intProductCode
                Debug.Print intIssueYear
                Debug.Print intPremiumDiscount
                Debug.Print intRemainingPOI

                intRow = 4
                Do While Not IsEmpty(Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intProductCode).Value)
                    strReadProductCode = Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intProductCode).Value
                    intReadIssueYear = Year(Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intIssueYear).Value)
                    dblPremiumDiscount = Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intPremiumDiscount).Value
                    dblRemainingPOI = Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intRemainingPOI).Value

                    For i = 1 To 8
                        If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                            'row of writing result
                            intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                            dblDac = dblPremiumDiscount * dblRemainingPOI * 0 'times zero for premium discount
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                                Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblDac
                        End If
                    Next
                    intRow = intRow + 1
                Loop
            End If

            Workbooks(strFile).Close SaveChanges:=False
        End If
        
        If Mid(strFile, 1, 29) = "Portfolio Inforce_Individual_" Then
            Set wb = Workbooks.Open(strFolder & strFile)
            
            '1222 0123.... 0124
            strIndividual = Mid(strFile, 30, 4)
            'column of writing result
            intColResult = Int(Mid(strIndividual, 1, 2)) + 12 * (Int(Mid(strIndividual, 3, 2)) - 22) - 9
            
            shtData = "Data IF"
            intProductCode = 70: intIssueYear = 70: intUD = 70: intPD = 70: intCommission = 70
            For i = 1 To 70
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Issued Year" Then intIssueYear = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Unearned Days" Then intUD = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Policy Duration (Days)" Then intPD = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Commission" Then intCommission = i
            Next
            Debug.Print intProductCode
            Debug.Print intIssueYear
            Debug.Print intUpr
            Debug.Print intRiUpr
            
            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value
                intReadIssueYear = Val(Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value)
                intUnearnedDays = Workbooks(strFile).Sheets(shtData).Cells(intRow, intUD).Value
                intPolicyDurationDays = Workbooks(strFile).Sheets(shtData).Cells(intRow, intPD).Value
                dblCommission = Workbooks(strFile).Sheets(shtData).Cells(intRow, intCommission).Value
                
                For i = 1 To 8
                    If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                        'row of writing result
                        intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                        dblDac = intUnearnedDays * dblCommission / intPolicyDurationDays
                        Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblDac
                    End If
                Next
                intRow = intRow + 1
            Loop
            
            Workbooks(strFile).Close SaveChanges:=False
        End If

        strFile = Dir
    Loop
End Sub

Public Sub PrepareIBNRGOC()
    Dim wb As Workbook
    Dim strFolder As String
    Dim strFile As String
    
    strWorkbook = ActiveWorkbook.Name
    strSheetResult = "Result"

    strFolder = "C:\Actuarial-BAU\IFRS17\"
    strFile = Dir(strFolder & "*.xls*")
        
    Do While strFile <> ""
        
        If Mid(strFile, 1, 24) = "Portfolio Inforce_Group_" Then
        'If Mid(strFile, 1, 7) = "Claims_" Then
            Set wb = Workbooks.Open(strFolder & strFile)

            'to check if SLC DS sheet exists
            sheetExists = False
            For Each ws In Workbooks(strFile).Worksheets
                If ws.Name = "SCL DS" Then
                    sheetExists = True
                    'MsgBox strFile
                    Exit For
                End If
            Next ws
            'MsgBox sheetExists

            'to check if sheet 3 month
            intIndex = 1
            For Each ws In Workbooks(strFile).Worksheets
                intIndex = intIndex + 1
                If Mid(ws.Name, 1, 7) = "Data EX" Then Exit For
            Next ws
            'MsgBox intIndex
            'Exit Do

            '1222 0123.... 0124
            strGroup = Mid(strFile, 25, 4)
            'column of writing result
            intColResult = Int(Mid(strGroup, 1, 2)) + 12 * (Int(Mid(strGroup, 3, 2)) - 22) - 9
            shtData = "Data IF"
            
            intProductCode = 70: intIssueYear = 70: intEarnedPremium = 70: intRiEarnedPremium = 70
            For i = 1 To 70
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Issue Year" Then intIssueYear = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Earned Premium" Then intEarnedPremium = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "RI Earned Premium" Then intRiEarnedPremium = i
            Next
            
            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value
                intReadIssueYear = Val(Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value)
                dblEarnedPremium = Workbooks(strFile).Sheets(shtData).Cells(intRow, intEarnedPremium).Value
                dblRiEarnedPremium = Workbooks(strFile).Sheets(shtData).Cells(intRow, intRiEarnedPremium).Value
                
                For i = 1 To 8
                    If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                        'row of writing result
                        intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                        'dblIBNR = dblEarnedPremium
                        Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblRiEarnedPremium 'dblEarnedPremium
                    End If
                Next
                intRow = intRow + 1
            Loop
            
            'Data EX <= 3 Months
            intProductCode = 70: intIssueYear = 70: intEarnedPremium = 70: intRiEarnedPremium = 70
            For i = 1 To 70
                If Workbooks(strFile).Sheets(intIndex).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(intIndex).Cells(2, i).Value = "Issue Year" Then intIssueYear = i
                If Workbooks(strFile).Sheets(intIndex).Cells(2, i).Value = "Earned Premium" Then intEarnedPremium = i
                If Workbooks(strFile).Sheets(intIndex).Cells(2, i).Value = "RI Earned Premium" Then intRiEarnedPremium = i
            Next

            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(intIndex).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(intIndex).Cells(intRow, intProductCode).Value
                intReadIssueYear = Val(Workbooks(strFile).Sheets(intIndex).Cells(intRow, intIssueYear).Value)
                dblEarnedPremium = Workbooks(strFile).Sheets(intIndex).Cells(intRow, intEarnedPremium).Value
                dblRiEarnedPremium = Workbooks(strFile).Sheets(intIndex).Cells(intRow, intRiEarnedPremium).Value

                For i = 1 To 8
                    If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                        'row of writing result
                        intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                        'dblIBNR = dblEarnedPremium
                        Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblRiEarnedPremium 'dblEarnedPremium
                    End If
                Next
                intRow = intRow + 1
            Loop
            
            If sheetExists Then
                shtDataSCL = "SCL DS"
                intProductCode = 70: intIssueYear = 70: intEarnedPremium = 70: intRiEarnedPremium = 70
                For i = 1 To 70
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "product_code" Then intProductCode = i
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "issue_date" Then intIssueYear = i
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "Earned Premium" Then intEarnedPremium = i
                    If Workbooks(strFile).Sheets(shtDataSCL).Cells(3, i).Value = "RI Earned Premium" Then intRiEarnedPremium = i
                Next
                
                intRow = 4
                Do While Not IsEmpty(Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intProductCode).Value)
                    strReadProductCode = Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intProductCode).Value
                    intReadIssueYear = Year(Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intIssueYear).Value)
                    dblEarnedPremium = Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intEarnedPremium).Value
                    dblRiEarnedPremium = Workbooks(strFile).Sheets(shtDataSCL).Cells(intRow, intRiEarnedPremium).Value
                    
                    For i = 1 To 8
                        If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                            'row of writing result
                            intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                            'dblIBNR = dblEarnedPremium
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                                Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblRiEarnedPremium 'dblEarnedPremium
                        End If
                    Next
                    intRow = intRow + 1
                Loop
            End If

            Workbooks(strFile).Close SaveChanges:=False
        End If
        
        If Mid(strFile, 1, 29) = "Portfolio Inforce_Individual_" Then
            Set wb = Workbooks.Open(strFolder & strFile)

            '1222 0123.... 0124
            strIndividual = Mid(strFile, 30, 4)
            'column of writing result
            intColResult = Int(Mid(strIndividual, 1, 2)) + 12 * (Int(Mid(strIndividual, 3, 2)) - 22) - 9

            shtData = "Data IF"
            
            intProductCode = 70: intIssueYear = 70: intEarnedPremium = 70: intRiEarnedPremium = 70
            For i = 1 To 70
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Issued Year" Then intIssueYear = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Earned Premium" Then intEarnedPremium = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "RI Earned Premium" Then intRiEarnedPremium = i
            Next

            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value
                intReadIssueYear = Val(Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value)
                dblEarnedPremium = Workbooks(strFile).Sheets(shtData).Cells(intRow, intEarnedPremium).Value
                dblRiEarnedPremium = Workbooks(strFile).Sheets(shtData).Cells(intRow, intRiEarnedPremium).Value
                
                For i = 1 To 8
                    If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                        'row of writing result
                        intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                        Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblRiEarnedPremium 'dblEarnedPremium
                    End If
                Next
                intRow = intRow + 1
            Loop

            Workbooks(strFile).Close SaveChanges:=False
        End If

        strFile = Dir
    Loop
End Sub

Public Sub PrepareIBNRLossRatioGOC()
    Dim wb As Workbook
    Dim strFolder As String
    Dim strFile As String
    
    strWorkbook = ActiveWorkbook.Name
    strSheetResult = "Result"

    strFolder = "C:\Actuarial-BAU\IFRS17\"
    strFile = Dir(strFolder & "*.xls*")

    Do While strFile <> ""
        
        If Mid(strFile, 1, 7) = "Claims_" Then
            Set wb = Workbooks.Open(strFolder & strFile)
                        
            '1222 0123.... 0124
            strClaim = Mid(strFile, 8, 4)
            'column of writing result
            intColResult = Int(Mid(strClaim, 1, 2)) + 12 * (Int(Mid(strClaim, 3, 2)) - 22) - 9
            shtData = "Index"
            
            intProductCode = 70: intScenario = 70: intLossRatio = 70: intNetLossRatio = 70
            For i = 1 To 11
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Scenario" Then intScenario = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Expected Loss Ratio (Gross)" Then intLossRatio = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Expected Loss Ratio (Net)" Then intNetLossRatio = i
            Next
            Debug.Print intProductCode
            Debug.Print intScenario
            Debug.Print intLossRatio
            
            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value
                strScenario = Workbooks(strFile).Sheets(shtData).Cells(intRow, intScenario).Value
                dblLossRatio = Workbooks(strFile).Sheets(shtData).Cells(intRow, intLossRatio).Value
                dblNetLossRatio = Workbooks(strFile).Sheets(shtData).Cells(intRow, intNetLossRatio).Value
                
                If strScenario = "Padded" Then
                    For i = 1 To 8
                        If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                            'row of writing result
                            For j = 2022 To 2024
                                intRowResult = (2025 - j) * 9 + i - 8
                                Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                                    dblNetLossRatio 'dblLossRatio
                            Next
                        End If
                    Next
                End If
                
                intRow = intRow + 1
            Loop
            

            Workbooks(strFile).Close SaveChanges:=False
        End If
        
        strFile = Dir
    Loop
End Sub


Public Sub PrepareIBNRClaimGOC()
    Dim wb As Workbook
    Dim strFolder As String
    Dim strFile As String
    
    strWorkbook = ActiveWorkbook.Name
    strSheetResult = "Result"

    strFolder = "C:\Actuarial-BAU\IFRS17\"
    strFile = Dir(strFolder & "*.xls*")
        
    Do While strFile <> ""
        
        If Mid(strFile, 1, 7) = "Claims_" Then
            Set wb = Workbooks.Open(strFolder & strFile)
                        
            '1222 0123.... 0124
            strClaim = Mid(strFile, 8, 4)
            'column of writing result
            intColResult = Int(Mid(strClaim, 1, 2)) + 12 * (Int(Mid(strClaim, 3, 2)) - 22) - 9
            shtData = "Claims"
            
            datValuationDate = Workbooks(strFile).Sheets("Index").Cells(2, 3).Value
            
            intProductCode = 70: intIssueYear = 70: intClaimStatus = 70: intReportDate = 70: intOSClaim = 70: intRiOSClaim = 70
            For i = 1 To 70
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Product Code" Then intProductCode = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Policy Effective Date" Then intIssueYear = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Claim Status" Then intClaimStatus = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Reported Period" Then intReportDate = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Claim Outstanding Reserve" Then intOSClaim = i
                If Workbooks(strFile).Sheets(shtData).Cells(2, i).Value = "Claim RI Outstanding Recovery" Then intRiOSClaim = i
            Next
            
            intRow = 3
            Do While Not IsEmpty(Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value)
                strReadProductCode = Workbooks(strFile).Sheets(shtData).Cells(intRow, intProductCode).Value
                If Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value = "" Then
                    intReadIssueYear = 2022
                Else
                    intReadIssueYear = Year(Workbooks(strFile).Sheets(shtData).Cells(intRow, intIssueYear).Value)
                End If
                datReportDate = Workbooks(strFile).Sheets(shtData).Cells(intRow, intReportDate).Value
                strClaimStatus = Workbooks(strFile).Sheets(shtData).Cells(intRow, intClaimStatus).Value
                dblReadOSClaim = Workbooks(strFile).Sheets(shtData).Cells(intRow, intOSClaim).Value
                dblReadRiOSClaim = Workbooks(strFile).Sheets(shtData).Cells(intRow, intRiOSClaim).Value
                
                If strClaimStatus <> "Rejected" And datReportDate <= datValuationDate Then
                    For i = 1 To 8
                        If Workbooks(strWorkbook).Sheets(strSheetResult).Cells(i + 1, 1).Value = strReadProductCode Then
                            'row of writing result
                            intRowResult = (2025 - intReadIssueYear) * 9 + i - 8
                            Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value = _
                                Workbooks(strWorkbook).Sheets(strSheetResult).Cells(intRowResult, intColResult).Value + dblReadRiOSClaim 'dblReadOSClaim ' ' ' ' '
                        End If
                    Next
                End If
                
                intRow = intRow + 1
            Loop
            

            Workbooks(strFile).Close SaveChanges:=False
        End If
        
        strFile = Dir
    Loop
End Sub

