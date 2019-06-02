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

