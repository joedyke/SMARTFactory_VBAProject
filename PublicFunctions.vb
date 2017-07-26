Public Function WeekDays(ByVal startDate As Date, ByVal endDate As Date) As Integer
   ' Returns the number of weekdays in the period from startDate
    ' to endDate not inclusive of start day. Returns blank if an error occurs.
    ' If your weekend days do not include Saturday and Sunday and
    ' do not total two per week in number, this function will
    ' require modification.
    On Error GoTo Weekdays_Error
    
    ' The number of weekend days per week.
    Const ncNumberOfWeekendDays As Integer = 2
    
    ' The number of days inclusive.
    Dim varDays As Variant
    
    ' The number of weekend days.
    Dim varWeekendDays As Variant
    
        
    ' Calculate the number of days not inclusive of start day
    varDays = DateDiff(Interval:="d", _
        date1:=startDate, _
        Date2:=endDate)
    
    
    ' Calculate the number of weekend days.
    varWeekendDays = (DateDiff(Interval:="ww", _
        date1:=startDate, _
        Date2:=endDate) _
        * ncNumberOfWeekendDays) _
        + IIf(DatePart(Interval:="w", _
        Date:=startDate) = vbSunday, 1, 0) _
        + IIf(DatePart(Interval:="w", _
        Date:=endDate) = vbSaturday, 1, 0)
    
    ' Calculate the number of weekdays.
    WeekDays = (varDays - varWeekendDays)
    
Weekdays_Exit:
    Exit Function
    
Weekdays_Error:
    WeekDays = ""
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, "Weekdays"
    Resume Weekdays_Exit
End Function

'==========================================================
' The DateAddW() function provides a workday substitute
' for DateAdd("w", number, date). This function performs
' error checking and ignores fractional Interval values.
' adds a number of weekdays to a particular date
'==========================================================
Function DateAddW(ByVal TheDate, ByVal Interval)
 
   Dim Weeks As Long, OddDays As Long, Temp As String
 
   If VarType(TheDate) <> 7 Or VarType(Interval) < 2 Or _
              VarType(Interval) > 5 Then
      DateAddW = TheDate
   ElseIf Interval = 0 Then
      DateAddW = TheDate
   ElseIf Interval > 0 Then
      Interval = Int(Interval)
 
   ' Make sure TheDate is a workday (round down).
 
      Temp = Format(TheDate, "ddd")
      If Temp = "Sun" Then
         TheDate = TheDate - 2
      ElseIf Temp = "Sat" Then
         TheDate = TheDate - 1
      End If
 
   ' Calculate Weeks and OddDays.
 
      Weeks = Int(Interval / 5)
      OddDays = Interval - (Weeks * 5)
      TheDate = TheDate + (Weeks * 7)
 
  ' Take OddDays weekend into account.
 
      If (DatePart("w", TheDate) + OddDays) > 6 Then
         TheDate = TheDate + OddDays + 2
      Else
         TheDate = TheDate + OddDays
      End If
 
      DateAddW = TheDate
    Else                         ' Interval is < 0
      Interval = Int(-Interval) ' Make positive & subtract later.
 
   ' Make sure TheDate is a workday (round up).
 
      Temp = Format(TheDate, "ddd")
      If Temp = "Sun" Then
         TheDate = TheDate + 1
      ElseIf Temp = "Sat" Then
         TheDate = TheDate + 2
      End If
 
   ' Calculate Weeks and OddDays.
 
      Weeks = Int(Interval / 5)
      OddDays = Interval - (Weeks * 5)
      TheDate = TheDate - (Weeks * 7)
 
   ' Take OddDays weekend into account.
 
      If (DatePart("w", TheDate) - OddDays) > 2 Then
         TheDate = TheDate - OddDays - 2
      Else
         TheDate = TheDate - OddDays
      End If
 
      DateAddW = TheDate
    End If
 
End Function
'**************  End of Code **************

Public Function DaysInMonth(Optional dtmDate As Date = 0) As Integer
    ' Return the number of days in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    DaysInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate) + 1, 1) - _
     DateSerial(Year(dtmDate), Month(dtmDate), 1)
End Function

'Add later
Function UpdateTimeStamp()
    
End Function

Function TestCallCreateCumulativePlan(DBpath)
    Dim obj As Object
    Set obj = CreateObject("Access.Application")
    With obj
        .OpenCurrentDatabase DBpath
        .Visible = False
        .Run "CreateCumulativePlan"
    End With
    
    obj.Quit
    Set obj = Nothing
End Function

Function CreateCumulativePlan()
    Dim cnn As Object 'connection object
    Dim rs As Object 'record set object
    Dim WorkingDays, totDays, DemandInput, ReqPerDay, _
        ReqPerDayRemainder, ReqPerDayNew, i, ExtraUnit, ReqDlvTot As Integer
    Dim sql, Program, ShortDate As String
    
    Set cnn = CreateObject("ADODB.Connection")
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = path_DB
        .Open
    End With
    
    'Delete records in tblCumulativePlan
    sql = "DELETE tblCumulativePlan.* FROM tblCumulativePlan;"
    cnn.Execute sql
    
    'Create sql code for query "qryMonthPlan" less the nz function
    sql = "SELECT tblDemandInput.Program, Sum([tblDemandInput].[PlnQTY]) AS PlanQTYTot" & _
          " FROM tblDemandInput" & _
          " GROUP BY tblDemandInput.Program" & _
          " HAVING (((tblDemandInput.Program) Not Like '*CCA' And (tblDemandInput.Program) Not Like '*SUB' And (tblDemandInput.Program)<>'HOSPITAL' And (tblDemandInput.Program)<>'TEAM'));"

    Set rs = cnn.Execute(sql)
    
    totDays = DaysInMonth
    
    
    'Calc number of working days in month (number of weekdays between two dates plus one)
    WorkingDays = WeekDays(DateSerial(Year(Date), Month(Date), 1), DateSerial(Year(Date), Month(Date), totDays)) + 1
    
    'For each program in qryMonthPlan...
    With rs
        If Not .BOF And Not .EOF Then
            While (Not .EOF)
                Program = rs.Fields("Program")
                
                'DemandInput, ReqPerDay, ReqPerDayRemainder
                DemandInput = rs.Fields("PlanQTYTot")
                If IsNull(DemandInput) Then
                    DemandInput = 0
                End If
                ReqPerDay = Int(DemandInput / WorkingDays)
                ReqPerDayRemainder = DemandInput Mod WorkingDays
                
                'For each day in the month...1 to totDays
                ReqDlvTot = 0 'set to zero
                For i = 1 To totDays
                    
                    'calc ShortDate
                    ShortDate = Month(Date) & "/" & (i) & "/" & Year(Date)
                    
                    'Enter required complete per day not including weekends
                    If Weekday(ShortDate) = "1" Or Weekday(ShortDate) = "7" Then
                        ReqPerDayNew = 0
                    Else
                        'we will add one additional unit every working day until the remainder is equal to zero
                        If ReqPerDayRemainder = 0 Then
                            ExtraUnit = 0
                        Else
                            ExtraUnit = 1
                            'subtract 1 from ShipPerDayRemainder
                            ReqPerDayRemainder = ReqPerDayRemainder - 1
                        End If
                        ReqPerDayNew = ReqPerDay + ExtraUnit
                    End If
                    ReqDlvTot = ReqDlvTot + ReqPerDayNew
                    'Send info to tblCumulativePlan
                    sql = "INSERT INTO tblCumulativePlan (shortdate, ReqDlvTot, Program ) " & _
                          "SELECT #" & ShortDate & "#, " & ReqDlvTot & ", '" & Program & "';"
                    cnn.Execute (sql)
               Next i
                .MoveNext
            Wend
        End If
    End With
    
    rs.Close
    cnn.Close
    Set cnn = Nothing
    Set rs = Nothing
    
End Function


'This function compacts a specified database.
'Note the database cannot be open for this function to be successful
'If the database is open an error will be stored and the routine will close
Public Sub CompactDb()
On Error GoTo CompactDb_Err
    Dim path2 As String
    Dim objAccess As Object 'database object (connection)
    Dim objDAO As Object
    
    path2 = Replace(path_DB, ".accde", "_temp.accde")
    
    Set objDAO = CreateObject("DAO.DBEngine.120")
    objDAO.CompactDataBase path_DB, path2
    
    Set objDAO = Nothing
    
    Kill path_DB
    Name path2 As path_DB
 
CompactDb_Exit:
    Set objDAO = Nothing
    Exit Sub

CompactDb_Err:
    Call LogErrorDesc(Error$, "CompactDB")
    Resume CompactDb_Exit

End Sub

'Logs the error description, time, and function or sub that the error occured in
Public Function LogErrorDesc(ErrTxt As String, ErrLoc As String)
    Dim ws  As Worksheet
    Dim element As Variant
    Dim row As Integer
    Set ws = Sheets("Error_Log")
               
    row = ws.Cells.Find(What:="*", _
                        After:=Range("A1"), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).row + 1
                
    
    ws.Cells(row, 1).value = Now
    ws.Cells(row, 2).value = ErrTxt
    ws.Cells(row, 3).value = ErrLoc
    
    Set ws = Nothing
End Function
