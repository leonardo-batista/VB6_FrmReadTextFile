Attribute VB_Name = "moduleSubrotine"
Option Explicit

Public Sub LoadConfigINI()

    Set gFileIni = New cFileIni
    
    gFileIni.Title = GetINISetting("SYSTEM", "TITLE", gFileIni.PathFile)
    gFileIni.System = GetINISetting("SYSTEM", "SYSTEM", gFileIni.PathFile)
    gFileIni.Campaign = GetINISetting("SYSTEM", "CAMPAIGN", gFileIni.PathFile)
    
    gFileIni.Header = GetINISetting("FILE", "HEADER", gFileIni.PathFile)
    gFileIni.DelimiterColumn = GetINISetting("FILE", "DELIMITER_COLUMN", gFileIni.PathFile)
    
    gFileIni.DataSource = GetINISetting("ODBC", "DATASOURCE", gFileIni.PathFile)
    gFileIni.Database = GetINISetting("ODBC", "DATABASE", gFileIni.PathFile)
    gFileIni.User = GetINISetting("ODBC", "USER", gFileIni.PathFile)
    gFileIni.Password = GetINISetting("ODBC", "PWD", gFileIni.PathFile)
    
    LogSystem "INFO", "LoadConfigINI", 0, "Read INI File was executed ;)"

End Sub

Public Sub GetInformation()

    gNameMachine = GetNameMachine
    gUserMachine = GetUserMachine
    TestConnection
    
End Sub

Public Sub SystemDirectory()
    On Error GoTo HandleError
    
    Dim strPath As String
    Dim strDirectory As String
    
    strPath = App.Path
    strDirectory = "LOG"
    
    'LOG
    If Dir(strPath & "\" & strDirectory, vbDirectory) = "" Then
        MkDir strPath & "\" & strDirectory
        LogSystem "INFO", "SystemDirectory", 0, "Directory LOG was created"
    End If
           
    strDirectory = "FILE"
           
    'FILE
    If Dir(strPath & "\" & strDirectory, vbDirectory) = "" Then
        MkDir strPath & "\" & strDirectory
        LogSystem "INFO", "SystemDirectory", 0, "Directory FILE was created"
    End If
    
    strDirectory = "INVALID"
    
    'INVALID
    If Dir(strPath & "\" & strDirectory, vbDirectory) = "" Then
        MkDir strPath & "\" & strDirectory
        LogSystem "INFO", "SystemDirectory", 0, "Directory INVALID was created"
    End If
           
    strDirectory = "LOADED"
    
    'LOADED
    If Dir(strPath & "\" & strDirectory, vbDirectory) = "" Then
        MkDir strPath & "\" & strDirectory
        LogSystem "INFO", "SystemDirectory", 0, "Directory LOADED was created"
    End If
           
HandleError:
    Debug.Print Err.Number & " " & Err.Description
    
End Sub

Public Sub LogSystem(levelError$, rotineName$, code$, message$)
    On Error GoTo HandleError

    Dim nUnit As Integer
    
    nUnit = FreeFile
    
    Open App.Path & "\LOG\" & App.EXEName & "_" & Format(Now, "yyyy-MM-dd") & ".log" For Append As nUnit
        Print #nUnit, "[START] " & Format(Now, "yyyy-MM-dd hh:nn:ss")
        Print #nUnit, levelError$ & " - Routine Performed: "; rotineName$
        Print #nUnit, "Code: " & code$ & " - Message: " & message$
        Print #nUnit, "[End]" & Chr(13)
    Close nUnit

HandleError:
    Debug.Print Err.Number & " " & Err.Description
    
End Sub

Public Sub TestConnection()
    On Error GoTo HandleError

        gConnectionDB.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=" & gFileIni.Database & ";uid=" & gFileIni.User & ";pwd=" & gFileIni.Password & ""
        gRecordsetDB.Open "SELECT CONVERT(VARCHAR(10), GETDATE(), 121) + ' ' + CONVERT(VARCHAR(8), GETDATE(), 108) AS DATE", gConnectionDB, adOpenStatic, adLockReadOnly
        
        If gRecordsetDB.RecordCount > 0 Then
            gDateAccess = gRecordsetDB.Fields!Date
        Else
            gDateAccess = "Problems with Database Connection !!!"
        End If

        gRecordsetDB.Close
        gConnectionDB.Close
        
        LogSystem "INFO", "TestConnection", 0, "Connection Test with Database was executed with SUCCESS, Date: " & gDateAccess
        
HandleError:
    Debug.Print Err.Number & " " & Err.Description
    
End Sub

