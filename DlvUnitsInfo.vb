Private pathMB52, pathZSD_SHP, pathZSD_SDEL As String ' file paths

'Initialize file paths
Public Sub InitializeFilePaths(ByVal path_MB52Input As String, _
            ByVal path_ZSD_SHPInput As String, _
            ByVal path_ZSD_SDELInput As String)
    pathMB52 = path_MB52Input
    pathZSD_SHP = path_ZSD_SHPInput
    pathZSD_SDEL = path_ZSD_SDELInput
End Sub


'This function parses .txt file "MB52" and sends that to the
'DB to create tblFinishedGoods
Public Function MB52TextToDB()
    Dim row, i As Integer
    Dim dictKeys() As String 'array used to store column headers in dictKeys
    Dim cnn As Object 'database object (connection)
    Dim sql, sql2 As String
    
    Set cnn = CreateObject("ADODB.Connection")
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = path_DB
        .Open
    End With
    
    
    'Change tblReadyForRefreshFlag
    sql = "UPDATE tblReadyForRefresh SET tblReadyForRefresh.ReadyForRefresh = False WHERE (((tblReadyForRefresh.ID)=1))"
    cnn.Execute sql
    
    
    'Clear table before sending data
    tblName = "tblFinishedGoods"
    sql = "DELETE " & tblName & ".* FROM " & tblName & ";"
    cnn.Execute sql
    
    path = pathMB52

    row = 1 'initiate row
    Open path For Input As #1        'open text file
    While Not EOF(1)                 'while the current line is not the end of file
        Line Input #1, LineText      'read the line of text
        Dim arr 'array variable used to split on the tab delimiter
        
        'If the linetext isn't a blank line then parse (don't parse a blank line)
        If Not LineText = "" And row > 1 Then
            LineText = Right(LineText, Len(LineText) - 1) 'cut off the leading tab delimiter
        
            LineText = Replace(LineText, ",", "") 'remove all commas
            LineText = Replace(LineText, "'", "") 'remove all apostrophes
            LineText = Replace(LineText, "%", "") 'remove all percent signs
            LineText = Replace(LineText, "#", "") 'remove all pound signs
            
            arr = Split(CStr(LineText), vbTab) 'split the LineText on tab
            
            'Start parsing at row 4 (when the headers start)
            If row = 2 Then
                LineText = Replace(LineText, " ", "") 'remove all spaces
                LineText = Replace(LineText, "/", "") 'remove all slash marks in the header
                LineText = Replace(LineText, ".", "") 'remove all periods
                arr = Split(CStr(LineText), vbTab)    'split the headers again (already split but we want the / removed so just do it again)
                
                ReDim dictKeys(LBound(arr) To UBound(arr))
                i = 0
                'Add column headers as keys in dictKeys
                For Each element In arr
                    dictKeys(i) = element
                    i = i + 1
                Next
                
                
            End If
            
            'Ready to write to Db
            If row > 2 Then
                
                sql = "INSERT INTO " & tblName & " ("
                sql2 = " SELECT "
                i = 0
                'send arr array to database.
                For Each element In arr
                    If Not dictKeys(i) = "" Then
                        If UBound(arr) = i Then
                            sql = sql & "[" & dictKeys(i) & "]) "
                            sql2 = sql2 & "'" & element & "'"
                        Else
                            sql = sql & "[" & dictKeys(i) & "], "
                            sql2 = sql2 & "'" & element & "', "
                        End If
                    End If
                    i = i + 1
                Next
                sql = sql & sql2
                cnn.Execute sql
                
                
                
            End If
            
        End If
        
        row = row + 1
    Wend
    Close #1
    
    'Change tblReadyForRefreshFlag
    sql = "UPDATE tblReadyForRefresh SET tblReadyForRefresh.ReadyForRefresh = True WHERE (((tblReadyForRefresh.ID)=1))"
    cnn.Execute sql
    
    'Close connection
    cnn.Close
    Set cnn = Nothing
    
End Function


'This function parses .txt file "ZSD_SHP" and sends that to the
'DB to create tblOTD
Public Function SHPTextToDb()
    Dim row, i As Integer
    Dim dictKeys() As String 'array used to store column headers in dictKeys
    Dim cnn As Object 'database object (connection)
    Dim sql, sql2 As String
    
    Set cnn = CreateObject("ADODB.Connection")
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = path_DB
        .Open
    End With
    
    'Change tblReadyForRefreshFlag
    sql = "UPDATE tblReadyForRefresh SET tblReadyForRefresh.ReadyForRefresh = False WHERE (((tblReadyForRefresh.ID)=1))"
    cnn.Execute sql
    
    
    'Clear table before sending data
    tblName = "tblOTD"
    sql = "DELETE " & tblName & ".* FROM " & tblName & ";"
    cnn.Execute sql
    
    path = pathZSD_SHP

    row = 1 'initiate row
    Open path For Input As #1        'open text file
    While Not EOF(1)                 'while the current line is not the end of file
        Line Input #1, LineText      'read the line of text
        Dim arr 'array variable used to split on the tab delimiter
        
        'If the linetext isn't a blank line then parse (don't parse a blank line)
        If Not LineText = "" And row > 4 Then
            LineText = Right(LineText, Len(LineText) - 1) 'cut off the leading tab delimiter
        
            LineText = Replace(LineText, ",", "") 'remove all commas
            LineText = Replace(LineText, "'", "") 'remove all apostrophes
            LineText = Replace(LineText, "%", "") 'remove all percent signs
            LineText = Replace(LineText, "#", "") 'remove all pound signs
            
            arr = Split(CStr(LineText), vbTab) 'split the LineText on tab
            
            'Start parsing at row 4 (when the headers start)
            If row = 5 Then
                LineText = Replace(LineText, " ", "") 'remove all spaces
                LineText = Replace(LineText, "/", "") 'remove all slash marks in the header
                LineText = Replace(LineText, ".", "") 'remove all periods
                arr = Split(CStr(LineText), vbTab)    'split the headers again (already split but we want the / removed so just do it again)
                
                ReDim dictKeys(LBound(arr) To UBound(arr))
                i = 0
                'Add column headers as keys in dictKeys
                For Each element In arr
                    dictKeys(i) = element
                    i = i + 1
                Next
                
                
            End If
            
            'Ready to write to Db
            If row > 5 Then
                
                sql = "INSERT INTO " & tblName & "("
                sql2 = " SELECT "
                i = 0
                'send arr array to database.
                For Each element In arr
                    If Not dictKeys(i) = "" Then
                        If UBound(arr) = i Then
                            sql = sql & "[" & dictKeys(i) & "]) "
                            sql2 = sql2 & "'" & element & "'"
                        Else
                            sql = sql & "[" & dictKeys(i) & "], "
                            sql2 = sql2 & "'" & element & "', "
                        End If
                    End If
                    i = i + 1
                Next
                sql = sql & sql2
                cnn.Execute sql
                
            End If
            
        End If
        
        row = row + 1
    Wend
    Close #1
    
    
    'Change tblReadyForRefreshFlag
    sql = "UPDATE tblReadyForRefresh SET tblReadyForRefresh.ReadyForRefresh = True WHERE (((tblReadyForRefresh.ID)=1))"
    cnn.Execute sql
    
    'Close connection
    cnn.Close
    Set cnn = Nothing
End Function

Public Function SDELTextToDb()
    Dim row, i As Integer
    Dim dictKeys() As String 'array used to store column headers in dictKeys
    Dim cnn As Object 'database object (connection)
    Dim sql, sql2 As String
    
    Set cnn = CreateObject("ADODB.Connection")
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = path_DB
        .Open
    End With
    
    'Change tblReadyForRefreshFlag
    sql = "UPDATE tblReadyForRefresh SET tblReadyForRefresh.ReadyForRefresh = False WHERE (((tblReadyForRefresh.ID)=1))"
    cnn.Execute sql
    
    'Clear table before sending data
    tblName = "tblShipments"
    sql = "DELETE " & tblName & ".* FROM " & tblName & ";"
    cnn.Execute sql
    
    path = pathZSD_SDEL

    row = 1 'initiate row
    Open path For Input As #1        'open text file
    While Not EOF(1)                 'while the current line is not the end of file
        Line Input #1, LineText      'read the line of text
        Dim arr 'array variable used to split on the tab delimiter
        
        'If the linetext isn't a blank line then parse (don't parse a blank line)
        If Not LineText = "" And row > 6 Then
            LineText = Right(LineText, Len(LineText) - 1) 'cut off the leading tab delimiter
        
            
            LineText = Replace(LineText, ",", "") 'remove all commas
            LineText = Replace(LineText, "'", "") 'remove all apostrophes
            LineText = Replace(LineText, "%", "") 'remove all percent signs
            LineText = Replace(LineText, "#", "") 'remove all pound signs
            
            arr = Split(CStr(LineText), vbTab) 'split the LineText on tab
            
            
            'Start parsing at row 4 (when the headers start)
            If row = 7 Then
                LineText = Replace(LineText, " ", "") 'remove all spaces
                LineText = Replace(LineText, "/", "") 'remove all slash marks in the header
                LineText = Replace(LineText, ".", "") 'remove all periods
                arr = Split(CStr(LineText), vbTab)    'split the headers again (already split but we want the / removed so just do it again)
                
                ReDim dictKeys(LBound(arr) To UBound(arr))
                i = 0
                'Add column headers as keys in dictKeys
                For Each element In arr
                    dictKeys(i) = element
                    i = i + 1
                Next
                
                
            End If
            
            'Ready to write to Db
            If row > 7 Then
                
                sql = "INSERT INTO " & tblName & "("
                sql2 = " SELECT "
                i = 0
                'send arr array to database.
                For Each element In arr
                    If Not dictKeys(i) = "" Then
                        If UBound(arr) = i Then
                            sql = sql & "[" & dictKeys(i) & "]) "
                            sql2 = sql2 & "'" & element & "'"
                        Else
                            sql = sql & "[" & dictKeys(i) & "], "
                            sql2 = sql2 & "'" & element & "', "
                        End If
                    End If
                    i = i + 1
                Next
                sql = sql & sql2
                cnn.Execute sql
                
            End If
            
        End If
        
        row = row + 1
    Wend
    Close #1
    
    'Change tblReadyForRefreshFlag
    sql = "UPDATE tblReadyForRefresh SET tblReadyForRefresh.ReadyForRefresh = True WHERE (((tblReadyForRefresh.ID)=1))"
    cnn.Execute sql
    
    'Close connection
    cnn.Close
    Set cnn = Nothing
End Function
