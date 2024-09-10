Attribute VB_Name = "GlobalVariable"
'Created by Leonardo Sembiring, dated 28 January 2024

'Purpose : To update global variables in variable_global.R and run result run_result.R

Public Sub UpdateR()
    Call UpdateGlobalVariable
    Call UpdateRunResult
End Sub

Public Sub UpdateGlobalVariable()
    strSheetMainVariable = "Main Variable"
    
    intRow = 7 'just change this row while you insert another row above
    strMainDirectory = Sheets(strSheetMainVariable).Cells(intRow, 2).Value 'main directory for reserving files
    strCurrentPeriod = Sheets(strSheetMainVariable).Cells(intRow + 1, 2).Value 'current period of valuation
    strPreviousPeriod = Sheets(strSheetMainVariable).Cells(intRow + 2, 2).Value 'previous period of valuation
    strRProgrammingLocation = Sheets(strSheetMainVariable).Cells(intRow + 4, 2).Value 'R files located
    
    'environment : 1 = inforce validation, 2 = valuation, 3 = movement validation 4 = movement, 5 = reporting
    strEnvironment = Sheets(strSheetMainVariable).Cells(intRow + 5, 2).Value
    'common id variable across databases
    strCommonDataId = Sheets(strSheetMainVariable).Cells(intRow + 8, 2).Value
    
    'name of R file to be updated
    strFileName = "variable_global"
    'full name of variable_globarl.R
    Filename = strRProgrammingLocation & "/" & strFileName & ".R"
    
    'file system object to manipulate text file
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(Filename, True)
    
    'write into variable_global.R
    a.WriteLine "# This program was created by Leonardo Sembiring (22 January 2024)"
    a.WriteLine "# For global variables"
    a.WriteLine
    a.WriteLine "#current and previous valuation period"
    a.WriteLine "current.valuation.period <- " & """" & strCurrentPeriod & """"
    a.WriteLine "previous.valuation.period <- " & """" & strPreviousPeriod & """"
    a.WriteLine
    a.WriteLine "#main folder of valuation working files"
    a.WriteLine "main.directory <-" & """" & strMainDirectory & """"
    a.WriteLine "r.directory <-" & """" & strRProgrammingLocation & """"
    a.WriteLine
    a.WriteLine "#run environment : 1 = inforce validation, 2 = valuation, 3 = movement validation 4 = movement, 5 = reporting"
    a.WriteLine "run.environment <- " & Trim(Str(strEnvironment))
    a.WriteLine
    a.WriteLine "#common data id across the dataframes"
    a.WriteLine "common.data.id <- " & "make.names('" & strCommonDataId & "')"
    a.WriteLine
    a.WriteLine "#**************************** table setup ****************"
    a.WriteLine
    a.WriteLine GenerateProductTable(2, 1, "df.product.info")
    a.WriteLine
    a.WriteLine GenerateProductTable(2, 28, "df.rates")
    a.WriteLine
    a.WriteLine "#selection of columns and replace by new column"
    
    'writing data inforce : R format
    arrCols = GenerateColsX(3, 12, True)
    a.WriteLine "df.inforce.subset.select <- make.names(c(" & arrCols(1) & "))" '
    'a.WriteLine "df.inforce.subset.select <- c(" & arrCols(1) & ")" '
    a.WriteLine "df.inforce.subset.column.name <- c(" & arrCols(2) & ")"
    a.WriteLine
    a.WriteLine "#column type of dataframe to read data-inforce-YYYYMM"
    a.WriteLine "df.inforce.cols <- cols(" & arrCols(3) & ")"
    a.WriteLine
    
    'writing data premium : R format
    arrCols = GenerateColsX(3, 18, False)
    a.WriteLine "df.premium.subset.select <- make.names(c(" & arrCols(1) & "))" '
    'a.WriteLine "df.premium.subset.select <- c(" & arrCols(1) & ")" '
    a.WriteLine "df.premium.subset.column.name <- c(" & arrCols(2) & ")"
    a.WriteLine
    a.WriteLine "#column type of dataframe to read data-premium-YYYYMM"
    a.WriteLine "df.premium.cols <- cols(" & arrCols(3) & ")"
    a.WriteLine
    
    'writing data reinsurance : R format
    arrCols = GenerateColsX(3, 21, False)
    'a.WriteLine "df.reinsurance.subset.select <- make.names(c(" & arrCols(1) & "))"
    a.WriteLine "df.reinsurance.subset.select <- c(" & arrCols(1) & ")"
    a.WriteLine "df.reinsurance.subset.column.name <- c(" & arrCols(2) & ")"
    a.WriteLine
    a.WriteLine "#column type of dataframe to read data-reinsurance-YYYYMM"
    a.WriteLine "df.reinsurance.cols <- cols(" & arrCols(3) & ")"
    a.WriteLine
    
    'writing data claim reserve : R format
    arrCols = GenerateColsX(3, 24, False)
    a.WriteLine "df.claim.reserve.subset.select <- make.names(c(" & arrCols(1) & "))"
    'a.WriteLine "df.claim.reserve.subset.select <- c(" & arrCols(1) & ")"
    a.WriteLine "df.claim.reserve.subset.column.name <- c(" & arrCols(2) & ")"
    a.WriteLine
    a.WriteLine "#column type of dataframe to read data-claim-reserve-YYYYMM"
    a.WriteLine "df.claim.reserve.cols <- cols(" & arrCols(3) & ")"
    a.WriteLine
    
    'writing data claim expense : R format
    arrCols = GenerateColsX(3, 27, False)
    a.WriteLine "df.claim.expense.subset.select <- make.names(c(" & arrCols(1) & "))"
    'a.WriteLine "df.claim.expense.subset.select <- c(" & arrCols(1) & ")"
    a.WriteLine "df.claim.expense.subset.column.name <- c(" & arrCols(2) & ")"
    a.WriteLine
    a.WriteLine "#column type of dataframe to read data-claim-expense-YYYYMM"
    a.WriteLine "df.claim.expense.cols <- cols(" & arrCols(3) & ")"
    a.WriteLine
    
    'reading data inforce : R format
    a.WriteLine "#column type of dataframe to read source inforce data"
    strCols = GenerateCols(3, 1) 'GWP table starting from row 3, column 1
    a.WriteLine "df.inforce.source.cols <- cols(" & strCols & ")"
    a.WriteLine
    
    'reading data premium : R format
    a.WriteLine "#column type of dataframe to read source premium data"
    strCols = GenerateCols(3, 3) 'Premium table starting from row 3, column 3
    a.WriteLine "df.premium.source.cols <- cols(" & strCols & ")"
    a.WriteLine
    
    'reading data reinsurance : R format
    a.WriteLine "#column type of dataframe to read source reinsurance data"
    strCols = GenerateCols(3, 5) 'Reinsurance table starting from row 3, column 5
    a.WriteLine "df.reinsurance.source.cols <- cols(" & strCols & ")"
    a.WriteLine
    
    'reading data claim reserve : R format
    a.WriteLine "#column type of dataframe to read source claim reserve data"
    strCols = GenerateCols(3, 7) 'Claim reserve table starting from row 3, column 7
    a.WriteLine "df.claim.reserve.source.cols <- cols(" & strCols & ")"
    a.WriteLine
    
    'reading data claim expense : R format
    a.WriteLine "#column type of dataframe to read source claim expense data"
    strCols = GenerateCols(3, 9) 'Claim expense table starting from row 3, column 9
    a.WriteLine "df.claim.expense.source.cols <- cols(" & strCols & ")"
    a.WriteLine
    
    a.Close
    
    'converting file from Excel CSV into UTF-18 in order to be able reading by R
    Call convertTxttoUTF(Replace(Filename, "/", "\"), Replace(Filename, "/", "\"))

End Sub

Public Sub UpdateRunResult()
    strSheetMainVariable = "Main Variable"
    
    intRow = 7 'just change this row while you insert another row above
    strRProgrammingLocation = Sheets(strSheetMainVariable).Cells(intRow + 4, 2).Value 'R files located
    
    'environment : 1 = inforce validation, 2 = valuation, 3 = movement validation 4 = movement, 5 = reporting
    strEnvironment = Sheets(strSheetMainVariable).Cells(intRow + 5, 2).Value
    
    'name of R file to be updated
    strFileName = "run_result"
    'full name of run_result.R
    Filename = strRProgrammingLocation & "/" & strFileName & ".R"
    
    'file system object to manipulate text file
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(Filename, True)
    
    'write into run_result.R
    a.WriteLine "# This program was created by Leonardo Sembiring (22 January 2024)"
    a.WriteLine "# For run result"
    a.WriteLine

    Select Case strEnvironment
        Case 1

        Case 2
            'write dataframe into tables
            a.WriteLine "#write into csv file : inforce, premium, reinsurance, claim reserve, claim expense"
            a.WriteLine "write.csv2(df.inforce, current.inforce.file.name, row.names = FALSE, quote = TRUE)"
            a.WriteLine "df.run.log <- write.run.log(paste('Data inforce created in ', current.inforce.file.name, sep = ''), df.run.log)"
            a.WriteLine
            a.WriteLine "write.csv2(df.premium, current.premium.file.name, row.names = FALSE, quote = TRUE)"
            a.WriteLine "df.run.log <- write.run.log(paste('Data premium created in ',current.premium.file.name, sep = ''), df.run.log)"
            a.WriteLine
            a.WriteLine "write.csv2(df.reinsurance, current.reinsurance.file.name, row.names = FALSE, quote = FALSE)"
            a.WriteLine "df.run.log <- write.run.log(paste('Data reinsurance created in ', current.reinsurance.file.name, sep = ''), df.run.log)"
            a.WriteLine
            a.WriteLine "write.csv2(df.claim.reserve, current.claim.reserve.file.name, row.names = FALSE, quote = TRUE)"
            a.WriteLine "df.run.log <- write.run.log(paste('Data claim reserve created in ', current.claim.reserve.file.name, sep = ''), df.run.log)"
            a.WriteLine
            a.WriteLine "write.csv2(df.claim.expense, current.claim.expense.file.name, row.names = FALSE, quote = TRUE)"
            a.WriteLine "df.run.log <- write.run.log(paste('Data claim expense created in ', current.claim.expense.file.name, sep = ''), df.run.log)"
            
        Case 3
        
        Case 4
            'write dataframe into tables
            a.WriteLine "#create movement file, filename format : data-movement-YYYYMM-to-YYYYMM"
            a.WriteLine "write.csv2(df.movement, movement.file.name, row.names = FALSE, quote = TRUE)"

            a.WriteLine
            a.WriteLine "write.csv2(df.inforce.current, current.inforce.file.name, row.names = FALSE, quote = TRUE)"
            a.WriteLine "df.run.log <- write.run.log(paste('Data current inforce updated in ', current.inforce.file.name, sep = ''), df.run.log)"

            a.WriteLine
            a.WriteLine "write.csv2(df.inforce.previous, previous.inforce.file.name, row.names = FALSE, quote = TRUE)"
            a.WriteLine "df.run.log <- write.run.log(paste('Data previous inforce updated in ', previous.inforce.file.name, sep = ''), df.run.log)"
        
        Case 5

            'open inforce files for both period : current and previous
            a.WriteLine "#open file current, previous, and movement"
            a.WriteLine "df.inforce.current <- read.csv2(current.inforce.file.name, check.names = FALSE)"
            a.WriteLine "df.inforce.previous <- read.csv2(previous.inforce.file.name, check.names = FALSE)"
            a.WriteLine
            
            'replacement titles of column in R format
            a.WriteLine "#replace column titles in R format"
            a.WriteLine "df.inforce.current.cols.name <- colnames(df.inforce.current)"
            a.WriteLine "colnames(df.inforce.current) <- make.names(colnames(df.inforce.current))"
            a.WriteLine
            a.WriteLine "df.inforce.previous.cols.name <- colnames(df.inforce.previous)"
            a.WriteLine "colnames(df.inforce.previous) <- make.names(colnames(df.inforce.previous))"
            a.WriteLine
            
            'starting of reporting
            intRow = 8
            intCol = 4
            Do While Not IsEmpty(Sheets(strSheetMainVariable).Cells(intRow, intCol).Value)
                intCheck = Sheets(strSheetMainVariable).Cells(intRow, intCol + 1).Value
                intRowReport = Sheets(strSheetMainVariable).Cells(intRow, intCol + 2).Value
                intColReport = Sheets(strSheetMainVariable).Cells(intRow, intCol + 3).Value
                
                If intCheck = 1 Then
                    a.WriteLine "#prepare group-by result"
                    strCols = GroupBy(intRowReport, intColReport)
                    For i = 1 To 6
                        a.WriteLine strCols(i)
                    Next
                    a.WriteLine
                End If
                intRow = intRow + 1
            Loop
                        
            'replace titles into original ones
            a.WriteLine "colnames(df.inforce.current) <- df.inforce.current.cols.name"
            a.WriteLine "colnames(df.inforce.previous) <- df.inforce.previous.cols.name"

    End Select
    
    a.Close
    
    'converting file from Excel CSV into UTF-18 in order to be able reading by R
    Call convertTxttoUTF(Replace(Filename, "/", "\"), Replace(Filename, "/", "\"))

End Sub

'simplify group-by report
Public Function GroupBy(ByVal intRow As Integer, ByVal intCol As Integer)
    Dim tempArr As Variant
    ReDim tempArr(1 To 6)

    strCols = GroupByReport(intRow, intCol)
    strCols(1) = Mid(strCols(1), 1, Len(strCols(1)) - 1)

    tempArr(1) = strCols(5) & " <- " & strCols(3) & " %>% dplyr::group_by(" & strCols(1) & ") %>% summarise("
    tempArr(2) = strCols(2)
    tempArr(3) = ")"
    tempArr(4) = "#save the result into file"
    tempArr(5) = "write.csv(" & strCols(5) & ", paste(current.valuation.directory, '/report/', '" & strCols(4) & "-', " & strCols(6) & ", '.csv'" & ", sep = '')" & ", row.names = FALSE)"
    tempArr(6) = ""
    
    GroupBy = tempArr
End Function

'to prepare group by report
Public Function GroupByReport(ByVal intRow As Integer, ByVal intCol As Integer)
    strSheetReference = "Reporting"
    
    Dim tempArr As Variant
    ReDim tempArr(1 To 6)
    
    strReportName = Sheets(strSheetReference).Cells(intRow, intCol + 1).Value 'name of report
    strCheckCurrent = Sheets(strSheetReference).Cells(intRow + 1, intCol + 1).Value 'current or previous source data
    
    intRow = intRow + 3 'starting to read list of fields
    strCols = ""
    strKeys = ""
    
    Do While Not IsEmpty(Sheets(strSheetReference).Cells(intRow, intCol).Value)
        strTitle = Sheets(strSheetReference).Cells(intRow, intCol).Value  'name of field in the result file
        strColumnName = Sheets(strSheetReference).Cells(intRow, intCol + 1).Value 'name of field from source
        strFunctionName = Sheets(strSheetReference).Cells(intRow, intCol + 2).Value 'group by function
        strCheck = Sheets(strSheetReference).Cells(intRow, intCol + 3).Value 'check fields to be included in the grouping
        strCheckKey = Sheets(strSheetReference).Cells(intRow, intCol + 4).Value 'check group-by fields
        
        If strCheck = 1 Then
            If strFunctionName = "n" Then
                strCols = strCols & "'" & strTitle & "' = " & _
                    strFunctionName & "()," & vbCrLf
            Else
                strCols = strCols & "'" & strTitle & "' = " & _
                    strFunctionName & "(" & Replace(strColumnName, " ", ".") & ")," & vbCrLf
            End If
        End If
        
        If strCheckKey = 1 Then
            strKeys = strKeys & Replace(strColumnName, " ", ".") & ","
        End If

        intRow = intRow + 1
    Loop
    
    tempArr(1) = strKeys 'list of key fields for group-by
    tempArr(2) = strCols 'list of source fields
    tempArr(3) = IIf(strCheckCurrent = 0, "df.inforce.current", "df.inforce.previous") 'source dataframe
    tempArr(4) = strReportName 'name of the report
    tempArr(5) = "df.report." & Trim(Str(intCol))
    tempArr(6) = IIf(strCheckCurrent = 0, "current.period", "previous.period") 'variable period
    
    GroupByReport = tempArr
End Function

'To generate variable cols in file variable_global.R
Public Function GenerateCols(ByVal intRow As Integer, ByVal intCol As Integer)
    strSheetReference = "Table"
    
    strCols = ""
    intCount = 1
    Do While Not IsEmpty(Sheets(strSheetReference).Cells(intRow, intCol).Value)
        strOriginalColumn = Sheets(strSheetReference).Cells(intRow, intCol).Value
        strColumnType = Sheets(strSheetReference).Cells(intRow, intCol + 1).Value
        
        Select Case strColumnType
            Case "N"
                strDescription = "col_double()"
            Case "D"
                strDescription = "col_date(format = " & """" & "%d-%m-%Y" & """" & ")"
            Case "D1"
                strDescription = "col_date(format = " & """" & "%Y-%m-%d %H:%M:%S" & """" & ")"
            Case "D2"
                strDescription = "col_date(format = " & """" & "%d/%m/%Y" & """" & ")"
            Case "D3"
                strDescription = "col_date(format = " & """" & "%Y-%m-%d" & """" & ")"
            Case "I"
                strDescription = "col_integer()"
            Case Else
                strDescription = "col_character()"
        End Select

        strCols = strCols & "'" & strOriginalColumn & "' = " & strDescription & "," & vbCrLf
        intRow = intRow + 1
    Loop
    
    GenerateCols = strCols
End Function

'To generate variable cols in file variable_global.R
Public Function GenerateColsX(ByVal intRow As Integer, ByVal intCol As Integer, ByVal bolAdditional As Boolean) As Variant
    strSheetReference = "Table"
    
    Dim tempArr As Variant
    ReDim tempArr(1 To 3)
    
    strReplace = ""
    strOriginal = ""
    strCols = ""
    intCount = 1
    Do While Not IsEmpty(Sheets(strSheetReference).Cells(intRow, intCol).Value)
        strReplaceColumn = Sheets(strSheetReference).Cells(intRow, intCol).Value
        strOriginalColumn = Sheets(strSheetReference).Cells(intRow, intCol + 1).Value
        If bolAdditional Then
            strAdditionalColumn = Sheets(strSheetReference).Cells(intRow, intCol + 4).Value
        Else
            strAdditionalColumn = ""
        End If
        strColumnType = Sheets(strSheetReference).Cells(intRow, intCol + 5).Value
        
        Select Case strColumnType
            Case "N"
                strDescription = "col_double()"
            Case "D"
                strDescription = "col_date(format = " & """" & "%d-%m-%Y" & """" & ")"
            Case "D1"
                strDescription = "col_date(format = " & """" & "%Y-%m-%d %H:%M:%S" & """" & ")"
            Case "D2"
                strDescription = "col_date(format = " & """" & "%d/%m/%Y" & """" & ")"
            Case "D3"
                strDescription = "col_date(format = " & """" & "%Y-%m-%d" & """" & ")"
            Case "I"
                strDescription = "col_integer()"
            Case Else
                strDescription = "col_character()"
        End Select

        If Trim(strOriginalColumn) <> "" Or Trim(strAdditionalColumn) <> "" Then
            If Trim(strAdditionalColumn) <> "" Then
                strOriginal = strOriginal & IIf(intCount = 1, "", ", ") & "'" & strAdditionalColumn & "'"
            Else
                strOriginal = strOriginal & IIf(intCount = 1, "", ", ") & "'" & strOriginalColumn & "'"
            End If
            strReplace = strReplace & IIf(intCount = 1, "", ", ") & "'" & strReplaceColumn & "'"
            strCols = strCols & "'" & strReplaceColumn & "' = " & strDescription & "," & vbCrLf

            intCount = intCount + 1
        End If
        intRow = intRow + 1
    Loop
    tempArr(1) = strOriginal
    tempArr(2) = strReplace
    tempArr(3) = strCols
    
    GenerateColsX = tempArr
End Function

'to change table into dataframe R
Public Function GenerateProductTable(ByVal intRow As Integer, ByVal intCol As Integer, ByVal strDataFrameName As String)
    strSheetReference = "Product"
    
    strAll = strDataFrameName & " <- data.frame(" & vbCrLf
    intTitleRow = intRow
    Do While Not IsEmpty(Sheets(strSheetReference).Cells(intTitleRow, intCol).Value)
            
        strTitle = Sheets(strSheetReference).Cells(intTitleRow, intCol).Value  'title
        strContent = Sheets(strSheetReference).Cells(intTitleRow + 1, intCol).Value 'test content string or not
        If IsNumeric(strContent) Then strQuote = "" Else strQuote = "'"
        
        intRow = intTitleRow
        strCols = "   '" & strTitle & "' = c("

        Do While Not IsEmpty(Sheets(strSheetReference).Cells(intRow + 1, intCol).Value)
            strContent = Sheets(strSheetReference).Cells(intRow + 1, intCol).Value
            'if decimal in comma, replace with dot
            If strQuote = "" And InStr(1, strContent, ",") Then strContent = Replace(strContent, ",", ".")
            
            strCols = strCols & strQuote & strContent & strQuote & ","
            intRow = intRow + 1
        Loop
        
        If IsEmpty(Sheets(strSheetReference).Cells(intTitleRow, intCol + 1).Value) Then
            strCols = Mid(strCols, 1, Len(strCols) - 1) & ")" & vbCrLf
        Else
            strCols = Mid(strCols, 1, Len(strCols) - 1) & ")," & vbCrLf
        End If
        
        strAll = strAll & strCols
        
        intCol = intCol + 1
    Loop
    strAll = Mid(strAll, 1, Len(strAll) - 1) & ")"
    
    GenerateProductTable = strAll
End Function
