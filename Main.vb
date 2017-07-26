Option Explicit

Global path_DB As Variant 'path to database needs to be global for SQL connection

Public Sub mainprocedure()
    Dim StartTime As Double
    StartTime = timer
    Dim SecondsElapsed, PauseTime, Start As Double
    Dim DataFolder, pathCOOISOps, pathZPPWIP, pathCOOISHeader, _
        pathCOOISHeaderOpen, pathCOOISHeaderDlv, _
        pathMB52, pathZSD_SHP, pathZSD_SDEL, pathCOOIS_KIT_DLV, dir As String
    Dim dictLastRefresh As Object '{TeamName: Time of last run}
    Dim LastRefresh As Date       'Last refresh variable
    Dim element As Variant
    Dim i As Integer
    
    
    Dim cnn As Object 'database object (connection)
    Dim sql As String
    
    
    
    Set cnn = CreateObject("ADODB.Connection")
    
    Dim TeamNames(1 To 2) As String
    TeamNames(1) = "HUMS"
    TeamNames(2) = "Computers"
    'TeamNames(3) = "MISSILES"
    
    Set dictLastRefresh = CreateObject("Scripting.Dictionary")
    
    PauseTime = 5 'seconds to pause between loops
    
    

    'Loop to exit if after 5:00 pm
    'loops over all data files and runs the VBA project if
    'the files are new
    'always runs at least once
    Do While timer < 61200 'exit loop after 5:00 pm
        For Each element In TeamNames
            'set DB path
            path_DB = Replace(ActiveWorkbook.path, "Administrator", "") & element & "\" & element & "Dashboard.accde"
            

            
            'Set DataFolder
            DataFolder = Replace(ActiveWorkbook.path, "Administrator", "") & element & "\Data\"
        
            'Set SAP export file paths
            pathCOOISOps = DataFolder & "COOIS_Operations.txt"
            pathZPPWIP = DataFolder & "ZSC_ZPPWIP_SN.txt"
            pathCOOISHeaderOpen = DataFolder & "COOIS_Headers_Open.txt"
            pathCOOISHeaderDlv = DataFolder & "COOIS_Headers_Dlv.txt"
            pathZSD_SHP = DataFolder & "ZSD_SHP.txt"
            pathZSD_SDEL = DataFolder & "ZSD_SDEL.txt"
            pathCOOIS_KIT_DLV = DataFolder & "COOIS_KIT_DLV.txt"
            pathMB52 = DataFolder & "MB52.txt"
            
            'Defaults to 12:00 00 AM if empty
            LastRefresh = dictLastRefresh(element)

            'if the date stamp of MB52.txt file (which is ran last in the SAP reports) is after
            'the timestamp that is saved in lastrefresh then parse and push data
            If FileDateTime(pathCOOISOps) > LastRefresh Then
                'Run parsing functions
                'Create COOISInfo Class Instance
                Dim COOISInfoInstance As New COOISInfoClass
                COOISInfoInstance.InitializeFilePaths pathCOOISOps, pathZPPWIP, pathCOOISHeaderOpen, pathCOOISHeaderDlv, pathCOOIS_KIT_DLV
                COOISInfoInstance.TextToDictCOOIS_Operations
                COOISInfoInstance.FindCurrentOp
                COOISInfoInstance.FindTimeRemaining
                COOISInfoInstance.TextToDictZSC_ZPPWIP_SNS
                COOISInfoInstance.TextAddHeaderOpen
                COOISInfoInstance.TextToDictDlvKit
                COOISInfoInstance.TextAddHeaderDlv
                COOISInfoInstance.UpdateDB
                Set COOISInfoInstance = Nothing
                
                'Create DlvUnitsInfo Class Instance
                Dim DlvUnitsInfoInstance As New DlvUnitsInfo
                DlvUnitsInfoInstance.InitializeFilePaths pathMB52, pathZSD_SHP, pathZSD_SDEL
                DlvUnitsInfoInstance.MB52TextToDB
                DlvUnitsInfoInstance.SHPTextToDb
                DlvUnitsInfoInstance.SDELTextToDb
                Set DlvUnitsInfoInstance = Nothing
                
                Call CreateCumulativePlan
                
                'Update time stamp (sql update query)
                With cnn
                    .Provider = "Microsoft.ACE.OLEDB.12.0"
                    .ConnectionString = path_DB
                    .Open
                End With
                
                'Update time stamp
                sql = "UPDATE tblLastRefresh SET tblLastRefresh.LastRefresh = #" & FileDateTime(pathCOOISOps) & _
                      "# WHERE (((tblLastRefresh.ID)=1))"
                cnn.Execute sql
                
                'Close connection
                cnn.Close

                'MsgBox "we are here LastRefresh = " & LastRefresh & " FileDateTime(pathMB52) = " & FileDateTime(pathMB52) & " team = " & element
                dictLastRefresh(element) = FileDateTime(pathCOOISOps)
            End If
            
            
            'MsgBox "outside loop"
            
            'Put a pause in the code
            Start = timer
            Do While timer < Start + PauseTime
                DoEvents 'yield to other processes
            Loop

        Next
    Loop
    
    'Compact each database at the end of each day
    'all databases should be closed.
    For Each element In TeamNames
            'set DB path
            path_DB = Replace(ActiveWorkbook.path, "Administrator", "") & element & "\" & element & "Dashboard.accde"
            Call CompactDb
    Next
    
    'Close last refresh object
    Set dictLastRefresh = Nothing
    Set cnn = Nothing
    
    SecondsElapsed = Round(timer - StartTime, 2)
    'MsgBox SecondsElapsed
End Sub

Public Sub Test2()
    Dim cnn As Object 'database object (connection)
    Dim sql As String
    
    Set cnn = CreateObject("ADODB.Connection")
    
    path_DB = "T:\_Shared_Workspace\Joe_Dyke\SMART Factory\Developer\HUMS\HUMSTestDashboard.accdb"
    
    'Update time stamp (sql update query)
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = path_DB
        .Open
    End With
    
    
    sql = "[qryCallCreateCumulativePlan]"
    
    
    cnn.Execute (sql)
    
    'Close connection
    cnn.Close
    
    'Close last refresh object

    Set cnn = Nothing
    
    
End Sub
