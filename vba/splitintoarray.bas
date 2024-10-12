'to split include parameter, ie. 1;3;5;4
Public Function SplitInclude(ByVal strInclude As String)
    Dim arrInclude() As String
    
    strDelimiter = ";"
    arrInclude = Split(strInclude, strDelimiter)
    
    SplitInclude = arrInclude
End Function
