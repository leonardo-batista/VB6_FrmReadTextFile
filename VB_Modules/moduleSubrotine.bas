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

End Sub

Public Sub TestConnection()
    On Error GoTo HandleError

        gConnectionDB.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=" & gFileIni.Database & ";uid=" & gFileIni.User & ";pwd=" & gFileIni.Password & ""
        gRecordsetDB.Open "SELECT CONVERT(VARCHAR(10), GETDATE(), 121) + ' ' + CONVERT(VARCHAR(8), GETDATE(), 108) AS DATE", gConnectionDB, adOpenStatic, adLockReadOnly
        
        If gRecordsetDB.RecordCount > 0 Then
            Dim valueTest As String
            valueTest = gRecordsetDB.Fields!Date
        End If
            
        gRecordsetDB.Close
        gConnectionDB.Close
        
HandleError:
    Debug.Print Err.Number & " " & Err.Description
    
End Sub
