'This is a test
Private dictCOOISInfo As Object 'master dictionary where SO is key and dictOpDetails is the value
                                         '{ShopOrderNumber:{dictOpDetails}}
Private path_COOISOperation, path_ZPPWIP, _
        path_COOISHeaderOpen, path_COOISHeaderDlv, _
        path_COOIS_KIT_DLV As String 'Member variable for the path to the coois
Private OpDetailKeys() As String 'array of value fields in opDetailArray (used as keys) (order, actfinishdate, etc)
Private HeaderDetailKeys() As String 'array of headerdetailkeys
Private HeaderDetailKeysDlv() As String 'array of headerdetailkeys
Private SOs As New Collection 'Collection of SO numbers
Private dictCurrentOps As Object 'dictionary that holds the current operation index number
                                    '{ShopOrderNumber:IndexOfCurrentOp in dictCOOISInfo}
Private dictTimeRemaining As Object 'dictionary that holds time remaining on order
                                   '{ShopOrderNumber:TimeRemaining}
Private dictSOSN As Object 'dict {ShopOrderNumber:SN}
Private dictCOOIS_KIT_DLV As Object 'dict {ShopOrderNumber:kitdate}
Private dictCOOISHeaderOpen As Object 'dictionary that holds COOIS info prior to sending to database
                                        '{ShopOrderNumberSN:{ShopOrderNumber:SOvalue, CurrentOp: Value,
                                                            'Actual Release: Value, PN: value, etc, Age: value}
Private dictCOOISHeaderDlv As Object 'dictionary that holds COOIS delivered orders info prior to sending to database
Private dictCOOISKitDate As Object 'dictionary that holds the kitting seq sign off date for open orders
                                    '{SO: KitSignOffDate}

'Initialize class with file paths
Private Sub Class_Initialize()
    path_COOISOperation = vbNullString
    path_ZPPWIP = vbNullString
End Sub

Public Sub InitializeFilePaths(ByVal path_COOISOpInput As String, _
            ByVal path_ZPPWIPInput As String, _
            ByVal path_COOISHeaderOpenInput As String, _
            ByVal path_COOISHeaderDlvInput As String, _
            ByVal path_COOIS_KIT_DLVInput As String)
    path_COOISOperation = path_COOISOpInput
    path_ZPPWIP = path_ZPPWIPInput
    path_COOISHeaderOpen = path_COOISHeaderOpenInput
    path_COOISHeaderDlv = path_COOISHeaderDlvInput
    path_COOIS_KIT_DLV = path_COOIS_KIT_DLVInput
End Sub

'Properties get values
Property Get getPath_COOISOperation() As String
    getPath_COOISOperation = path_COOISOperation
End Property

'Properties get values
Property Get getPath_ZPPWIP() As String
    getPath_ZPPWIP = path_ZPPWIP
End Property

'Methods
Public Function TextToDictCOOIS_Operations()
    'This function reads the SAP tab delimited COOIS operations .txt file (xls file saved as .txt)
    'and parses that information into a dictionary of a dictionary of arrays
    'as explained below
    
    'Created by Joseph Dyke
    
    Dim path As String 'local path
    Dim LineText As String 'variable for line from text file
    Dim ActStartPos As New Collection 'collection of position indexes that use the header ActStart
    Dim ActFinishPos As New Collection 'collection of position indexes that use the header ActFinish
    Dim UserFieldPos As New Collection 'collection of position indexes that use the header UserFieldPos
    Dim element As Variant 'variable to represent elements in arrays, collections, or dictionaries
    Dim row, i As Integer     'counters for the current row or iterating in array in the text file
    Dim SOIndex, WrkCntrIndex As Integer   'variable used to determine the index of SO number in arr
    Dim dictOpDetails As Object 'dictionary for the operation details of a particular order (will be inside of dictCOOISInfo)
                                '{OperAct:[value1, value2, etc], Operationshorttext:[value1, value2], etc}
    Dim OpDetailArray() As String 'variable array used to store detail infor for each order (overwritten every loop)
    Dim OpDetailKeysReady As Boolean 'flag used later
        OpDetailKeysReady = False
    Set dictOpDetails = CreateObject("Scripting.Dictionary")
    Set dictCOOISInfo = CreateObject("Scripting.Dictionary")
    
    
    path = path_COOISOperation
    
    row = 1 'initiate row
    Open path For Input As #1        'open text file
    While Not EOF(1)                 'while the current line is not the end of file
        Line Input #1, LineText      'read the line of text
        Dim arr 'array variable used to split on the tab delimiter
        
        'If the linetext isn't a blank line then parse (don't parse a blank line)
        If Not LineText = "" And row > 3 Then
            LineText = Right(LineText, Len(LineText) - 1) 'cut off the leading tab delimiter
        
            LineText = Replace(LineText, ",", "") 'remove all commas
            LineText = Replace(LineText, "'", "") 'remove all apostrophes
            LineText = Replace(LineText, "%", "") 'remove all percent signs
            LineText = Replace(LineText, "#", "") 'remove all pound signs
            
            arr = Split(CStr(LineText), vbTab) 'split the LineText on tab
            
            'Start parsing at row 3 (when the headers start)
            If row = 4 Then
                LineText = Replace(LineText, " ", "") 'remove all spaces
                LineText = Replace(LineText, "/", "") 'remove all slash marks in the header
                LineText = Replace(LineText, ".", "") 'remove all periods
                arr = Split(CStr(LineText), vbTab)    'split the headers again (already split but we want the / removed so just do it again)
                
                ReDim OpDetailKeys(1 To (UBound(arr) + 1)) As String 'set array size of OpDetailKeys to the number of elements in arr
                OpDetailKeys = arr
                
            End If
            
            'We want to use the headers as keys in our dictionary but there
            'are repeated headers "Actstart", "Actfinish", and "userfield". There are two per
            'because one is a date and one is a time or a stamp (userfield). We will look at the first
            'row of values and append "date" and "time" to the correct key
            If row = 6 Then  'row 6 is the first line of data
                'First determine where ActStart and Actfinish are in the array
                'And store them in a collection
                i = 0
                For Each element In OpDetailKeys
                    If element = "Actstart" Or element = "ActlStart" Then
                        ActStartPos.Add i 'stores the indices of the ActStart header
                    End If
                    If element = "Actfinish" Or element = "Actlfinish" Then
                        ActFinishPos.Add i 'stores the indices of the ActFinish header
                    End If
                    If element = "Userfield" Then
                        UserFieldPos.Add i 'stores the indices of the ActFinish header
                    End If
                    i = i + 1
                Next
                
                'Now determine which one is date and which one is time based
                'on the data in row 6. Append date or time to key depending
                'on if the data contains / (date) or : (time)
                For i = 1 To ActFinishPos.Count
                    If InStr(arr(ActFinishPos(i)), ":") <> 0 Then    'times contain :
                        OpDetailKeys(ActStartPos(i)) = OpDetailKeys(ActStartPos(i)) & "Time"  'append time to key
                        OpDetailKeys(ActFinishPos(i)) = OpDetailKeys(ActFinishPos(i)) & "Time"  'append time to key
                    Else
                        OpDetailKeys(ActStartPos(i)) = OpDetailKeys(ActStartPos(i)) & "Date"  'append date to key
                        OpDetailKeys(ActFinishPos(i)) = OpDetailKeys(ActFinishPos(i)) & "Date"  'append date to key
                    End If
                Next i
                
                For i = 1 To UserFieldPos.Count
                    If InStr(arr(UserFieldPos(i)), "/") <> 0 Then    'one userfield is a date so contains /
                        OpDetailKeys(UserFieldPos(i)) = OpDetailKeys(UserFieldPos(i)) & "Date"  'append date to key
                    Else
                        OpDetailKeys(UserFieldPos(i)) = OpDetailKeys(UserFieldPos(i)) & "Stamp"
                    End If
                Next i
                
                'flag to indicate the OpDetailKeys are ready for parsing
                OpDetailKeysReady = True
                
                'Detemine and store the index of the SO number
                i = 0
                For Each element In OpDetailKeys
                    If element = "Order" Then
                        SOIndex = i
                    End If
                Next
                
                'Loop over OpDetailKeys and initialize the dictOpDetails and dictCOOISInfo
                If OpDetailKeysReady = True Then
                    i = 0
                    For Each element In OpDetailKeys
                        ReDim OpDetailArray(1) As String
                        On Error Resume Next
                        OpDetailArray(1) = arr(i)
                        If Err.Number = 9 Then
                            OpDetailArray(1) = "empty"
                            Err.Clear
                        End If
                        On Error GoTo 0
                        
                        'Create dictionary of arrays
                        dictOpDetails.Add element, OpDetailArray
                        '{Order: [SO number1, SO number2, etc], OperAct: [OperAct1, OperAct2, etc], etc}
                        
                        i = i + 1
                    Next
                    
                    'Put dictOpDetails in dictCOOISInfo array the key is the shop order number
                    dictCOOISInfo.Add arr(SOIndex), dictOpDetails
                    '{OrderNumber1: dictOpDetails, OrderNumber2: dictOpDetails
                    
                    
                    
                End If
                
               
            End If
            
            If row > 6 Then 'we are now past the headers
                'Create a collection of unique SO numbers
                On Error Resume Next
                SOs.Add arr(SOIndex), arr(SOIndex)
                On Error GoTo 0
                
                'look and see if we are on a new shop order
                If dictCOOISInfo.Exists(arr(SOIndex)) Then  'if this is the same shop order add values to the array
                    i = 0
                    For Each element In OpDetailKeys
                        'redim lbound ubound
                        ReDim OpDetailArray(LBound(dictCOOISInfo(arr(SOIndex))(element)) To UBound(dictCOOISInfo(arr(SOIndex))(element)))
                        'set it equal to current array in dictionary
                        OpDetailArray = dictCOOISInfo(arr(SOIndex))(element)
                        
                        'increase array size by one
                        ReDim Preserve OpDetailArray(LBound(dictCOOISInfo(arr(SOIndex))(element)) To UBound(dictCOOISInfo(arr(SOIndex))(element)) + 1)
                        'Add latest value to array
                        On Error Resume Next
                        OpDetailArray(UBound(dictCOOISInfo(arr(SOIndex))(element)) + 1) = arr(i)
                        If Err.Number = 9 Then
                            OpDetailArray(UBound(dictCOOISInfo(arr(SOIndex))(element)) + 1) = "empty"
                            Err.Clear
                        End If
                        On Error GoTo 0
                        
                        
                        'overwrite array in dictOpDetails with new array with one more value
                        dictOpDetails(element) = OpDetailArray
                        'add value to array in dictOpDetails in dictCOOISInfo
                        i = i + 1
                    Next
                    
                    'Put dictOpDetails in dictCOOISInfo array the key is the shop order number
                    dictCOOISInfo.Remove arr(SOIndex)
                    dictCOOISInfo.Add arr(SOIndex), dictOpDetails
                    '{OrderNumber1: dictOpDetails, OrderNumber2: dictOpDetails
                    
                Else 'otherwise re-initialize the array
                    Set dictOpDetails = Nothing
                    Set dictOpDetails = CreateObject("Scripting.Dictionary")
                    i = 0
                    For Each element In OpDetailKeys
                        ReDim OpDetailArray(1) As String 'clears array
                        On Error Resume Next
                        OpDetailArray(1) = arr(i)
                        If Err.Number = 9 Then
                            OpDetailArray(1) = "empty"
                            Err.Clear
                        End If
                        On Error GoTo 0
                        
                        
                        'Update
                        dictOpDetails(element) = OpDetailArray
                        '{Order: [SO number1, SO number2, etc], OperAct: [OperAct1, OperAct2, etc], etc}
                        
                        i = i + 1
                    Next
                    
                    'Put dictOpDetails in dictCOOISInfo array the key is the shop order number
                    dictCOOISInfo.Add arr(SOIndex), dictOpDetails
                    '{OrderNumber1: dictOpDetails, OrderNumber2: dictOpDetails
                End If

            End If
            
            
        End If
        
        row = row + 1
    Wend
    Close #1
    
    
    
    Set dictOpDetails = Nothing
    Set ActStartPos = Nothing
    Set ActFinishPos = Nothing
    Set UserFieldPos = Nothing
End Function

'This function loops over the sequences in
'dictCOOISInfo and determines which operation
'is the current operation based on sign offs
'and d-stamps
'It also finds and stores the kitting sequence signoff date.
'Note if OpIndex number is 0 then all sequences are signed off
Public Function FindCurrentOp()
    Dim currentstamp, previousstamp, prepreviousstamp, kitdate As String
    Dim tempArray As Variant 'Used to store array of sequences
    Dim foundcurrentop As Boolean
    Dim i, j As Integer
    
    Set dictCurrentOps = CreateObject("Scripting.Dictionary")
    Set dictCOOISKitDate = CreateObject("Scripting.Dictionary")
    
    i = 0
    foundcurrentop = False
    For Each Outerkey In dictCOOISInfo.Keys 'loop over shop orders
        For Each element In dictCOOISInfo(Outerkey)("Workcntr")
            If Left(element, 3) = "KIT" Then
                kitdate = dictCOOISInfo(Outerkey)("UserfieldDate")(j)
                'Some routings have more than 1 kitting operation. If that is the case only store the first one
                On Error Resume Next
                dictCOOISKitDate.Add Outerkey, kitdate 'store index of kitting
                Err.Clear
                On Error GoTo 0
            End If
            j = j + 1
        Next
    
        For Each element In dictCOOISInfo(Outerkey)("UserfieldStamp") 'loop over array of sequences
        
            'Determine kitting signoff date
        
            If foundcurrentop = False Then 'if we havent found the current op continue
                If i = 0 Then 'if the first element is empty
                    'do nothing
                Else
                    currentstamp = element
                    If currentstamp = "" Then 'if the currentstamp is empty then check if the previous stamp
                                              'was a D-stamp
                        If InStr(previousstamp, "DIS") <> 0 Then 'If previous seq was d-stamped append "Hospital"
                            ReDim tempArray(LBound(dictCOOISInfo(Outerkey)("Operationshorttext")) To UBound(dictCOOISInfo(Outerkey)("Operationshorttext")))
                            tempArray = dictCOOISInfo(Outerkey)("Operationshorttext")
                            tempArray(i) = tempArray(i) & "_HOSPITAL"
                            dictCOOISInfo(Outerkey)("Operationshorttext") = tempArray
                        End If
                        
                        'store index of current op
                        dictCurrentOps.Add Outerkey, i
                        
                        
                        foundcurrentop = True

                    End If
                    
                    
                End If
                previousstamp = element
                i = i + 1
            End If
        Next
        If foundcurrentop = False Then
                'there are not any unsigned operations
                'store index of current 0 is unused in array so we will add "last operation" to is
                dictCurrentOps.Add Outerkey, 0
                ReDim tempArray(LBound(dictCOOISInfo(Outerkey)("Operationshorttext")) To UBound(dictCOOISInfo(Outerkey)("Operationshorttext")))
                tempArray = dictCOOISInfo(Outerkey)("Operationshorttext")
                tempArray(0) = "LAST OPERATION"
                dictCOOISInfo(Outerkey)("Operationshorttext") = tempArray
        End If
        foundcurrentop = False
        i = 0
        j = 0
    Next
    
End Function

'This function finds the amount of time remaining on the order
'For the sequences after current op
'Initializes dictTimeRemaining
Public Function FindTimeRemaining()
    Dim totaltime As Double
    Dim temptime1, temptime2, temptime3 As String
    Set dictTimeRemaining = CreateObject("Scripting.Dictionary")
    
    For Each key In dictCurrentOps.Keys 'for each shop order in dictCurrentOps.Keys
        
        If dictCurrentOps(key) = 0 Then
            totaltime = 0
        Else
            
            'Loop over array of dictCOOISInfo starting at current op index
            For i = dictCurrentOps(key) To UBound(dictCOOISInfo(key)("Setup"))
                temptime1 = dictCOOISInfo(key)("Setup")(i)
                temptime2 = dictCOOISInfo(key)("Duration")(i)
                temptime3 = dictCOOISInfo(key)("Queuetime")(i)
                If temptime1 = "" Then
                    temptime1 = 0
                End If
                If temptime2 = "" Then
                    temptime2 = 0
                End If
                If temptime3 = "" Then
                    temptime3 = 0
                End If
                'add up totaltime use labor, setup, and queuetime. Omit wait time for now
                'note unit of queue time is hours not minutes so *60
                totaltime = totaltime + CDbl(temptime1) + CDbl(temptime2) + CDbl(temptime3) * 60
            Next
        End If
        dictTimeRemaining.Add key, totaltime
        
        totaltime = 0
    Next
    
End Function

'This function parses the ZSC_ZPPWIP_SN txt file and initializes the
'dictSOSN with the following info
'{ShopOrderNumber:SN}

Public Function TextToDictZSC_ZPPWIP_SNS()
    Dim row, i, SOIndex, SNIndex As Integer
    Dim SO, SN, SOSN As String
    Dim dictKeys() As String 'array used to store headers index of value in dictKeys
                             'is the same as index for that value in arr
    Dim SNs As New Collection 'Collection of serial numbers if there are multiple
    Set dictSOSN = CreateObject("Scripting.Dictionary")
    
    path = path_ZPPWIP

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
                    If element = "Order" Then
                        SOIndex = i
                    End If
                    If element = "Serialnumber" Then
                        SNIndex = i
                    End If
                    i = i + 1
                Next
                
                
            End If
            
            
            If row > 7 Then
                'Extract SO and SN
                SO = arr(SOIndex)
                SN = arr(SNIndex)
                
                'try to add SN to SNs collection (if that SO already exists)
                If dictSOSN.Exists(SO) = True Then
                    Set SNs = dictSOSN(SO)
                    SNs.Add SN
                    Set dictSOSN(SO) = SNs
                Else
                    Set SNs = Nothing
                    Set SNs = New Collection
                    SNs.Add SN
                    dictSOSN.Add SO, SNs
                End If
                
                
                
                
            End If
            
        End If
        
        row = row + 1
    Wend
    Close #1
    
End Function

Public Function TextAddHeaderOpen()
    'This function reads the SAP tab delimited COOIS headers .txt file (xls file saved as .txt)
    'and parses that information into a dictionary of a dictionaries
    'as explained below
    
    'Created by Joseph Dyke
    
    Dim path As String 'local path
    Dim LineText As String 'variable for line from text file
    Dim element As Variant 'variable to represent elements in arrays, collections, or dictionaries
    Dim row, i As Integer     'counters for the current row or iterating in array in the text file
    Dim SOIndex As Integer   'variable used to determine the index of SO number in arr
    Dim ActStartPos As New Collection 'collection of position indexes that use the header ActStart
    Dim ActFinishPos As New Collection 'collection of position indexes that use the header ActFinish
    Dim dictHeaderDetails As Object 'dictionary for the operation details of a particular order (will be inside of dictCOOISInfo)
                                '{OperAct:[value1, value2, etc], Operationshorttext:[value1, value2], etc}
    Set dictHeaderDetails = Nothing
    Dim tempdictHeaderDetails As Object
    Dim firstpass As Boolean
    firstpass = True
    Set dictCOOISHeaderOpen = CreateObject("Scripting.Dictionary")
    
    path = path_COOISHeaderOpen

    row = 1 'initiate row
    Open path For Input As #1        'open text file
    While Not EOF(1)                 'while the current line is not the end of file
        Line Input #1, LineText      'read the line of text
        Dim arr 'array variable used to split on the tab delimiter
        
        
        
        'If the linetext isn't a blank line then parse (don't parse a blank line)
        If Not LineText = "" And row > 3 Then
            LineText = Right(LineText, Len(LineText) - 1) 'cut off the leading tab delimiter
        
            LineText = Replace(LineText, ",", "") 'remove all commas
            LineText = Replace(LineText, "'", "") 'remove all apostrophes
            LineText = Replace(LineText, "%", "") 'remove all percent signs
            LineText = Replace(LineText, "#", "") 'remove all pound signs
            
            arr = Split(CStr(LineText), vbTab) 'split the LineText on tab
            
            'Start parsing at row 4 (when the headers start)
            If row = 4 Then
                LineText = Replace(LineText, " ", "") 'remove all spaces
                LineText = Replace(LineText, "/", "") 'remove all slash marks in the header
                LineText = Replace(LineText, ".", "") 'remove all periods
                arr = Split(CStr(LineText), vbTab)    'split the headers again (already split but we want the / removed so just do it again)
                
                ReDim HeaderDetailKeys(1 To (UBound(arr) + 1)) As String 'set array size to the number of elements in arr
                HeaderDetailKeys = arr
                
                
                'We want to use the headers as keys in our dictionary but there
                'are repeated headers "Release" and "Release".
                'One is the planned release date and the other is the actual release date
                'the first date is always the Actual release date (controled via SAP export scripts)
                i = 0
                For Each element In HeaderDetailKeys
    
                    If element = "Release" Then
                        If firstpass = True Then
                            HeaderDetailKeys(i) = "ActRelease"
                            firstpass = False
                        Else
                            HeaderDetailKeys(i) = "PlanRelease"
                        End If
                    End If
                    i = i + 1
                Next
                
                
            End If

            'We want to use the headers as keys in our dictionary but there
            'are repeated headers "Actstart", "Actfinish", and "userfield". There are two per
            'because one is a date and one is a time or a stamp (userfield). We will look at the first
            'row of values and append "date" and "time" to the correct key
            If row = 6 Then 'first line of data
                'First determine where ActStart and Actfinish are in the array
                'And store them in a collection
                i = 0
                For Each element In HeaderDetailKeys
                    If element = "ActStart" Or element = "ActlStart" Then
                        ActStartPos.Add i 'stores the indices of the ActStart header
                    End If
                    If element = "Actfinish" Or element = "ActlStart" Then
                        ActFinishPos.Add i 'stores the indices of the ActFinish header
                    End If
                    i = i + 1
                Next
                
                'Now determine which one is date and which one is time based
                'on the data in row 6. Append date or time to key depending
                'on if the data contains / (date) or : (time)
                'Note: the actual finish date will be empty for all open orders
                'so the SAP export cannot have the ActFinishDate value at the end of a row (last column)
                'otherwise the parsing breaks
                

                For i = 1 To ActFinishPos.Count
                    If InStr(arr(ActFinishPos(i)), ":") <> 0 Then    'times contain :
                        HeaderDetailKeys(ActStartPos(i)) = HeaderDetailKeys(ActStartPos(i)) & "Time"  'append time to key
                        HeaderDetailKeys(ActFinishPos(i)) = HeaderDetailKeys(ActFinishPos(i)) & "Time"  'append time to key
                    Else 'it is date which is empty if there is no date
                        HeaderDetailKeys(ActStartPos(i)) = HeaderDetailKeys(ActStartPos(i)) & "Date"  'append date to key
                        HeaderDetailKeys(ActFinishPos(i)) = HeaderDetailKeys(ActFinishPos(i)) & "Date"  'append date to key
                    End If
                Next i
                
                
                'Detemine and store the index of the SO number
                i = 0
                For Each element In HeaderDetailKeys
                    If element = "Order" Then
                        SOIndex = i
                    End If
                    i = i + 1
                Next

                
            End If
            
            
            
            If row > 5 Then 'row 6 is the first line of data
                Set dictHeaderDetails = CreateObject("Scripting.Dictionary")
                
                'Loop over OpDetailKeys and initialize the dictOpDetails and dictCOOISInfo
                i = 0
                For Each element In HeaderDetailKeys
                    'Create dictionary of arrays
                    dictHeaderDetails.Add element, arr(i)
                    '{Order: value, Material: value, Order Type: value, etc}
                    i = i + 1
                Next
                
                'Dimension variables to improve readability
                Dim SO, SN, SOSN, CurrentOpText As String
                Dim timeremaining, SOage As Double
                SO = dictHeaderDetails(HeaderDetailKeys(SOIndex))
                      
                'get time remaining
                timeremaining = dictTimeRemaining(SO) 'defaults to 0 if empty
                dictHeaderDetails.Add "TimeRemaining", timeremaining
                
                'get current operation
                If Not IsEmpty(dictCurrentOps(SO)) Then
                    CurrentOpText = dictCOOISInfo(SO)("Operationshorttext")(dictCurrentOps(SO))
                Else
                    CurrentOpText = "NoCurrentOpInfo"
                End If
                dictHeaderDetails.Add "Operationshorttext", CurrentOpText
               
                'get kit signoff date
                Dim kitdate As String
                kitdate = dictCOOISKitDate(SO)
                dictHeaderDetails.Add "KitDate", kitdate
               
                'calc age, dev to plan release, MfgFinishDate of SO
                Dim age, ReleaseDev, MfgFinishDate As String 'store age as string (age of Shop orders in days)
                If Not dictHeaderDetails("ActRelease") = "" And Not dictHeaderDetails("PlanRelease") = "" Then
                    age = CStr(Date - CDate(dictHeaderDetails("ActRelease")))
                    ReleaseDev = CStr(WeekDays(CDate(dictHeaderDetails("PlanRelease")), CDate(dictHeaderDetails("ActRelease"))))
                    MfgFinishDate = CStr(DateAddW(CDate(dictHeaderDetails("Basicfin")), CDbl(ReleaseDev)))
                Else
                    age = ""
                    ReleaseDev = ""
                    MfgFinishDate = CStr(dictHeaderDetails("Basicfin"))
                End If
                dictHeaderDetails.Add "Age", age
                dictHeaderDetails.Add "ReleaseDev", ReleaseDev
                dictHeaderDetails.Add "MfgFinishDate", MfgFinishDate
                
                
                
                'get SN and SOSN
                If IsEmpty(dictSOSN(SO)) Then
                    SN = "Empty"
                    SOSN = SO & "_" & SN
                    
                    dictHeaderDetails.Add "SN", SN
                    
                    'Put dictHeaderDetails in dictCOOISHeaderOpen array the key is the shop order number
                    dictCOOISHeaderOpen.Add SOSN, dictHeaderDetails
                    '{OrderNumber1-SN1: dictHeaderDetails, OrderNumber2-SN2: dictOpDetails}
                Else
                    For Each element In dictSOSN(SO)
                        Set tempdictHeaderDetails = Nothing
                        Set tempdictHeaderDetails = CreateObject("Scripting.Dictionary")
                        
                        
                        SN = element
                        If SN = "" Then
                            SN = "Empty"
                        End If
                        
                        SOSN = SO & "_" & SN
                        
                        dictHeaderDetails("SN") = SN
                
                        For Each key In dictHeaderDetails.Keys
                            tempdictHeaderDetails.Add key, dictHeaderDetails(key)
                        Next
                        'Put dictHeaderDetails in dictCOOISHeaderOpen array the key is the shop order number
                        dictCOOISHeaderOpen.Add SOSN, tempdictHeaderDetails
                        '{OrderNumber1-SN1: dictHeaderDetails, OrderNumber2-SN2: dictOpDetails}

                    Next
                End If
                
                
            End If
            
        End If
        
        row = row + 1
    Wend
    Close #1
    

    
    'Store added keys
    ReDim HeaderDetailKeys(LBound(HeaderDetailKeys) To dictHeaderDetails.Count)
    i = 0
    For Each key In dictHeaderDetails.Keys
        HeaderDetailKeys(i) = key
        i = i + 1
    Next
    
    Set tempdictHeaderDetails = Nothing
    Set dictHeaderDetails = Nothing
    Set ActStartPos = Nothing
    Set ActFinishPos = Nothing
End Function

Public Function TextAddHeaderDlv()
    'This function reads the SAP tab delimited COOIS headers .txt file (xls file saved as .txt)
    'and parses that information into a dictionary of a dictionaries
    'as explained below
    
    'Created by Joseph Dyke
    
    Dim path As String 'local path
    Dim LineText As String 'variable for line from text file
    Dim element As Variant 'variable to represent elements in arrays, collections, or dictionaries
    Dim row, i As Integer     'counters for the current row or iterating in array in the text file
    Dim SOIndex As Integer   'variable used to determine the index of SO number in arr
    Dim ActStartPos As New Collection 'collection of position indexes that use the header ActStart
    Dim ActFinishPos As New Collection 'collection of position indexes that use the header ActFinish
    Dim dictHeaderDetails As Object 'dictionary for the operation details of a particular order (will be inside of dictCOOISInfo)
                                '{OperAct:[value1, value2, etc], Operationshorttext:[value1, value2], etc}
    Set dictHeaderDetails = Nothing
    Dim firstpass As Boolean
    firstpass = True
    Set dictCOOISHeaderDlv = CreateObject("Scripting.Dictionary")
    
    path = path_COOISHeaderDlv
    
    row = 1 'initiate row
    Open path For Input As #1        'open text file
    While Not EOF(1)                 'while the current line is not the end of file
        Line Input #1, LineText      'read the line of text
        Dim arr 'array variable used to split on the tab delimiter
        
        
        
        'If the linetext isn't a blank line then parse (don't parse a blank line)
        If Not LineText = "" And row > 3 Then
            LineText = Right(LineText, Len(LineText) - 1) 'cut off the leading tab delimiter
        
            LineText = Replace(LineText, ",", "") 'remove all commas
            LineText = Replace(LineText, "'", "") 'remove all apostrophes
            LineText = Replace(LineText, "%", "") 'remove all percent signs
            LineText = Replace(LineText, "#", "") 'remove all pound signs
            
            arr = Split(CStr(LineText), vbTab) 'split the LineText on tab
            
            'Start parsing at row 4 (when the headers start)
            If row = 4 Then
                LineText = Replace(LineText, " ", "") 'remove all spaces
                LineText = Replace(LineText, "/", "") 'remove all slash marks in the header
                LineText = Replace(LineText, ".", "") 'remove all periods
                arr = Split(CStr(LineText), vbTab)    'split the headers again (already split but we want the / removed so just do it again)
                
                ReDim HeaderDetailKeysDlv(1 To (UBound(arr) + 1)) As String 'set array size to the number of elements in arr
                HeaderDetailKeysDlv = arr
                
                
                'We want to use the headers as keys in our dictionary but there
                'are repeated headers "Release" and "Release".
                'One is the planned release date and the other is the actual release date
                'the first date is always the actual release date (controled via SAP export scripts)
                i = 0
                For Each element In HeaderDetailKeysDlv
                    If element = "Release" Then
                        If firstpass = True Then
                            HeaderDetailKeysDlv(i) = "ActRelease"
                            firstpass = False
                        Else
                            HeaderDetailKeysDlv(i) = "PlanRelease"
                        End If
                    End If
                    i = i + 1
                Next
                
                
            End If

            'We want to use the headers as keys in our dictionary but there
            'are repeated headers "Actstart", "Actfinish", and "userfield". There are two per
            'because one is a date and one is a time or a stamp (userfield). We will look at the first
            'row of values and append "date" and "time" to the correct key
            If row = 6 Then 'first line of data
                'First determine where ActStart and Actfinish are in the array
                'And store them in a collection
                i = 0
                For Each element In HeaderDetailKeysDlv
                    If element = "ActStart" Then
                        ActStartPos.Add i 'stores the indices of the ActStart header
                    End If
                    If element = "Actfinish" Then
                        ActFinishPos.Add i 'stores the indices of the ActFinish header
                    End If
                    i = i + 1
                Next
                
                'Now determine which one is date and which one is time based
                'on the data in row 6. Append date or time to key depending
                'on if the data contains / (date) or : (time)
                'Note: the actual finish date will be empty for all open orders
                'so the SAP export cannot have the ActFinishDate value at the end of a row (last column)
                'otherwise the parsing breaks
                For i = 1 To ActFinishPos.Count
                    If InStr(arr(ActFinishPos(i)), ":") <> 0 Then    'times contain :
                        HeaderDetailKeysDlv(ActStartPos(i)) = HeaderDetailKeysDlv(ActStartPos(i)) & "Time"  'append time to key
                        HeaderDetailKeysDlv(ActFinishPos(i)) = HeaderDetailKeysDlv(ActFinishPos(i)) & "Time"  'append time to key
                    Else 'it is date which is empty if there is no date
                        HeaderDetailKeysDlv(ActStartPos(i)) = HeaderDetailKeysDlv(ActStartPos(i)) & "Date"  'append date to key
                        HeaderDetailKeysDlv(ActFinishPos(i)) = HeaderDetailKeysDlv(ActFinishPos(i)) & "Date"  'append date to key
                    End If
                Next i
                
                
                'Detemine and store the index of the SO number
                i = 0
                For Each element In HeaderDetailKeysDlv
                    If element = "Order" Then
                        SOIndex = i
                    End If
                    i = i + 1
                Next

                
            End If
            
            
            
            If row > 5 Then 'row 6 is the first line of data
                Set dictHeaderDetails = CreateObject("Scripting.Dictionary")
                
                'Loop over OpDetailKeys and initialize the dictOpDetails and dictCOOISInfo
                i = 0
                For Each element In HeaderDetailKeysDlv
                    'Create dictionary of arrays
                    dictHeaderDetails.Add element, arr(i)
                    '{Order: value, Material: value, Order Type: value, etc}
                    i = i + 1
                Next
                
                'Dimension variables to improve readability
                Dim SO, SN, SOSN, CurrentOpText As String
                Dim timeremaining, SOage As Double
                SO = dictHeaderDetails(HeaderDetailKeysDlv(SOIndex))
               
                'calc age, dev to plan release, MfgFinishDate of SO
                Dim age, ReleaseDev, MfgFinishDate As String 'store age as string (age of Shop orders in days)
                If Not dictHeaderDetails("ActRelease") = "" Then
                    age = CStr(Date - CDate(dictHeaderDetails("ActRelease")))
                    ReleaseDev = CStr(WeekDays(CDate(dictHeaderDetails("PlanRelease")), CDate(dictHeaderDetails("ActRelease"))))
                    MfgFinishDate = CStr(DateAddW(CDate(dictHeaderDetails("Basicfin")), CDbl(ReleaseDev)))
                Else 'if order isn't released (i.e. if it is a zero hour or something was done wrong put empty strings in
                    age = ""
                    ReleaseDev = ""
                    MfgFinishDate = CStr(dictHeaderDetails("Basicfin"))
                End If
                dictHeaderDetails.Add "Age", age
                dictHeaderDetails.Add "ReleaseDev", ReleaseDev
                dictHeaderDetails.Add "MfgFinishDate", MfgFinishDate
                
                'Add kitting signoff date
                kitdate = dictCOOIS_KIT_DLV(SO)
                dictHeaderDetails.Add "KitDate", kitdate
                
                'Put dictHeaderDetails in dictCOOISHeaderOpen array the key is the shop order number
                dictCOOISHeaderDlv.Add SO, dictHeaderDetails
                '{OrderNumber1: dictHeaderDetails1, OrderNumber2: dictHeaderDetails}
            End If
            
        End If
        
        row = row + 1
    Wend
    Close #1
    
    
    

    
    'Store added keys
    ReDim HeaderDetailKeysDlv(LBound(HeaderDetailKeysDlv) To dictHeaderDetails.Count)
    i = 0
    For Each key In dictHeaderDetails.Keys
        HeaderDetailKeysDlv(i) = key
        i = i + 1
    Next
    
    Set dictHeaderDetails = Nothing
    Set ActStartPos = Nothing
    Set ActFinishPos = Nothing
End Function

'This function parses "COOIS_KIT_DLV.txt" and
'creates the "dictCOOIS_KIT_DLV" dictionarity {Order: kitdate}
Public Function TextToDictDlvKit()
    Dim row, i, SOIndex, kitdateIndex As Integer
    Dim SO, kitdate As String
    Dim dictKeys() As String 'array used to store headers index of value in dictKeys
                             'is the same as index for that value in arr
    Dim SNs As New Collection 'Collection of serial numbers if there are multiple
    Set dictCOOIS_KIT_DLV = CreateObject("Scripting.Dictionary")
    
    path = path_COOIS_KIT_DLV
    

    row = 1 'initiate row
    Open path For Input As #1        'open text file
    While Not EOF(1)                 'while the current line is not the end of file
        Line Input #1, LineText      'read the line of text
        Dim arr 'array variable used to split on the tab delimiter
        
        'If the linetext isn't a blank line then parse (don't parse a blank line)
        If Not LineText = "" And row > 3 Then
            
            LineText = Right(LineText, Len(LineText) - 1) 'cut off the leading tab delimiter
        
            LineText = Replace(LineText, ",", "") 'remove all commas
            LineText = Replace(LineText, "'", "") 'remove all apostrophes
            LineText = Replace(LineText, "%", "") 'remove all percent signs
            LineText = Replace(LineText, "#", "") 'remove all pound signs
            
            arr = Split(CStr(LineText), vbTab) 'split the LineText on tab
            
            
            'Start parsing at row 4 (when the headers start)
            If row = 4 Then
                LineText = Replace(LineText, " ", "") 'remove all spaces
                LineText = Replace(LineText, "/", "") 'remove all slash marks in the header
                LineText = Replace(LineText, ".", "") 'remove all periods
                arr = Split(CStr(LineText), vbTab)    'split the headers again (already split but we want the / removed so just do it again)
                
                ReDim dictKeys(LBound(arr) To UBound(arr))
                
                i = 0
                'Add column headers as keys in dictKeys
                For Each element In arr
                    dictKeys(i) = element
                    If element = "Order" Then
                        SOIndex = i
                    End If
                    If element = "Userfield" Then
                        kitdateIndex = i
                    End If
                    i = i + 1
                Next
                
                
            End If
            
            
            If row > 4 And Not arr(1) = "No Data" Then
                'Extract SO and SN
                SO = arr(SOIndex)
                kitdate = arr(kitdateIndex)
                
                'the routing may have multiple kitting sequences
                'we only want the first one
                On Error Resume Next
                dictCOOIS_KIT_DLV.Add SO, kitdate
                Err.Clear
                On Error GoTo 0
                
            End If
            
        End If
        
        row = row + 1
    Wend
    Close #1
End Function

'This function sends the dictionary created in the above TextToDict function to the
'workbook to aid development
Public Function COOISInfoToWS(ByVal sheetname As String, _
                          ByVal rowstart As Integer, ByVal colstart As Integer)
    Dim ws  As Worksheet
    Dim key As Variant
    Dim key2 As Variant
    Dim i, j As Integer
    'Set Dict = CreateObject("Scripting.Dictionary")
    Set ws = Sheets(sheetname)
                       
    i = rowstart
    ws.Cells(i - 1, colstart).value = "COOISInfo Key Outer"
    ws.Cells(i - 1, colstart + 1).value = "COOISInfo Key Inner"
    ws.Cells(i - 1, colstart + 2).value = "Value"
    For Each key In dictCOOISInfo.Keys
        ws.Cells(i, colstart).value = key
        
        For Each key2 In dictCOOISInfo(key).Keys
            ws.Cells(i, colstart + 1).value = key2
            For j = 1 To UBound(dictCOOISInfo(key)(key2))
                ws.Cells(i, colstart + 1 + j).value = dictCOOISInfo(key)(key2)(j)
                
            Next j
            i = i + 1 'increment cell row
        Next
        
    Next
    
    Set ws = Nothing
End Function



'This function writes an array to the worksheet (1D) returns value and index
'used to aid development
Public Function OpDetailKeysToWS(ByVal sheetname As String, _
                          ByVal rowstart As Integer, ByVal colstart As Integer)
    Dim ws  As Worksheet
    Dim element As Variant
    Dim i, j As Integer
    Set ws = Sheets(sheetname)
                       
    i = rowstart
    j = 0
    ws.Cells(i - 1, colstart).value = "Op Detail Keys"
    ws.Cells(i - 1, colstart + 1).value = "Index"
    For Each element In OpDetailKeys
        ws.Cells(i, colstart).value = element
        ws.Cells(i, colstart + 1).value = j
        i = i + 1 'increment cell row
        j = j + 1 'increment index
    Next
    
    Set ws = Nothing
End Function

'This function writes an array to the worksheet (1D) returns value and index
'used to aid development
Public Function SONumsToWS(ByVal sheetname As String, _
                          ByVal rowstart As Integer, ByVal colstart As Integer)
    Dim ws  As Worksheet
    Dim element As Variant
    Dim i, j As Integer
    Set ws = Sheets(sheetname)
                       
    i = rowstart
    j = 0
    ws.Cells(i - 1, colstart).value = "Shop Orders"
    ws.Cells(i - 1, colstart + 1).value = "Index"
    For Each element In SOs
        ws.Cells(i, colstart).value = element
        ws.Cells(i, colstart + 1).value = j
        i = i + 1 'increment cell row
        j = j + 1 'increment index
    Next
    
    Set ws = Nothing
End Function

'This function writes an array to the worksheet (1D) returns value and index
'used to aid development
Public Function dictToWS(ByVal sheetname As String, _
                          ByVal rowstart As Integer, ByVal colstart As Integer, ByVal keyvaluename As String)
    Dim ws  As Worksheet
    Dim element As Variant
    Dim i, j As Integer
    Set ws = Sheets(sheetname)
                       
    i = rowstart
    j = 0
    ws.Cells(i - 1, colstart).value = "Shop Order"
    ws.Cells(i - 1, colstart + 1).value = keyvaluename
    For Each key In dictCurrentOps.Keys
        ws.Cells(i, colstart).value = key
        ws.Cells(i, colstart + 1).value = dictCurrentOps(key)
        i = i + 1 'increment cell row
        j = j + 1 'increment index
    Next
    
    Set ws = Nothing
End Function
'This function writes an array to the worksheet (1D) returns value and index
'used to aid development
Public Function dictToWS2(ByVal sheetname As String, _
                          ByVal rowstart As Integer, ByVal colstart As Integer, ByVal keyvaluename As String)
    Dim ws  As Worksheet
    Dim element As Variant
    Dim i, j As Integer
    Set ws = Sheets(sheetname)
                       
    i = rowstart
    j = 0
    ws.Cells(i - 1, colstart).value = "Shop Order"
    ws.Cells(i - 1, colstart + 1).value = keyvaluename
    For Each key In dictTimeRemaining.Keys
        ws.Cells(i, colstart).value = key
        ws.Cells(i, colstart + 1).value = dictTimeRemaining(key)
        i = i + 1 'increment cell row
        j = j + 1 'increment index
    Next
    
    Set ws = Nothing
End Function

Public Function checkCurrentOpIndex()
    For Each key In dictCurrentOps.Keys
        MsgBox dictCOOISInfo(key)("Operationshorttext")(dictCurrentOps(key))
    Next
End Function

Public Function UpdateDB()
    Dim cnn As Object 'database object (connection)
    Dim sql, sql2, tblName As String
    
    Set cnn = CreateObject("ADODB.Connection")
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = path_DB
        .Open
    End With
    
    'Change tblReadyForRefreshFlag
    sql = "UPDATE tblReadyForRefresh SET tblReadyForRefresh.ReadyForRefresh = False WHERE (((tblReadyForRefresh.ID)=1))"
    cnn.Execute sql
    
    
    'tblSOs Loop over SOs and send to access
    tblName = "tblSOs"
    sql = "DELETE " & tblName & ".* FROM " & tblName & ";"
    cnn.Execute sql
    For Each key In dictCurrentOps.Keys
        sql = "INSERT INTO " & tblName & "([Order]) SELECT '" & key & "'"
        cnn.Execute sql
        sql = ""
    Next
    
    'Send header detail keys
    tblName = "tblHeaderDetailKeys"
    sql = "DELETE " & tblName & ".* FROM " & tblName & ";"
    cnn.Execute sql
    For Each element In HeaderDetailKeys
        If Not element = "" Then
            sql = "INSERT INTO " & tblName & "([HeaderDetailKeys]) SELECT '" & element & "'"
            cnn.Execute sql
        End If
        sql = ""
    Next
    
    'Send OpDetailKeys
    tblName = "tblOpDetailKeys"
    sql = "DELETE " & tblName & ".* FROM " & tblName & ";"
    cnn.Execute sql
    For Each element In OpDetailKeys
        If Not element = "" Then
            sql = "INSERT INTO " & tblName & "([OpDetailKeys]) SELECT '" & element & "'"
            cnn.Execute sql
        End If
        sql = ""
    Next
    
    
    'Send SNs
    tblName = "tbldictSOSN"
    sql = "DELETE " & tblName & ".* FROM " & tblName & ";"
    cnn.Execute sql
    For Each key In dictSOSN.Keys

        'Could be multiple serial numbers per SO
        If Not IsEmpty(dictSOSN(key)) Then
            For Each element In dictSOSN(key)
                If Not element = "" Then
                    sql = "INSERT INTO " & tblName & "([Order], [SN]) SELECT '" & key & "', '" & element & "'"
                    cnn.Execute sql
                End If
            Next
        End If
    Next
    
   
    'send open order headerinfo to tblRawDataOpen
    tblName = "tblRawDataOpen"
    sql = "DELETE " & tblName & ".* FROM " & tblName & ";"
    cnn.Execute sql
    For Each key In dictCOOISHeaderOpen.Keys
        sql = "INSERT INTO " & tblName & "("
        sql2 = " SELECT "
        i = 1
        For Each key2 In HeaderDetailKeys
            If Not key2 = "" Then
                If UBound(HeaderDetailKeys) = i Then
                    sql = sql & "[" & key2 & "]) "
                    sql2 = sql2 & "'" & dictCOOISHeaderOpen(key)(key2) & "'"
                Else
                    sql = sql & "[" & key2 & "], "
                    sql2 = sql2 & "'" & dictCOOISHeaderOpen(key)(key2) & "', "
                    
                End If
                i = i + 1
            End If
        Next
        sql = sql & sql2
        cnn.Execute sql
    Next
    
    'send dlv order headerinfo to tblRawDataDlv
    tblName = "tblRawDataDlv"
    sql = "DELETE " & tblName & ".* FROM " & tblName & ";"
    cnn.Execute sql
    For Each key In dictCOOISHeaderDlv.Keys
        sql = "INSERT INTO " & tblName & "("
        sql2 = " SELECT "
        i = 1
        For Each key2 In HeaderDetailKeysDlv
            If Not key2 = "" Then
                If UBound(HeaderDetailKeysDlv) = i Then
                    sql = sql & "[" & key2 & "]) "
                    sql2 = sql2 & "'" & dictCOOISHeaderDlv(key)(key2) & "'"
                Else
                    sql = sql & "[" & key2 & "], "
                    sql2 = sql2 & "'" & dictCOOISHeaderDlv(key)(key2) & "', "
                    
                End If
                i = i + 1
            End If
        Next
        sql = sql & sql2
        cnn.Execute sql
    Next
    
    'Change tblReadyForRefreshFlag
    sql = "UPDATE tblReadyForRefresh SET tblReadyForRefresh.ReadyForRefresh = True WHERE (((tblReadyForRefresh.ID)=1))"
    cnn.Execute sql
    
    'cnn.Open cnnstring
    cnn.Close
    
    Set cnn = Nothing
End Function

