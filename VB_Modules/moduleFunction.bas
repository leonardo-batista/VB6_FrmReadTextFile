Attribute VB_Name = "moduleFunction"
Option Explicit
'Get the INI Setting from the File
Public Function GetINISetting(ByVal sHeading As String, ByVal sKey As String, sINIFileName) As String
    On Error GoTo HandleError
    
    Const cparmLen = 255
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim lLength As Long
    lLength = GetPrivateProfileString(sHeading, sKey _
            , sDefault, sReturn, cparmLen, sINIFileName)
    GetINISetting = Mid(sReturn, 1, lLength)

HandleError:
    Call LogSystem("ERROR", "GetINISetting", Err.Number, Err.Description)
End Function

Public Function GetNameMachine() As String
    On Error GoTo HandleError
    
    Dim strBuffer As String
    Dim strAns As Long
    
    strBuffer = Space$(255)
    strAns = GetComputerName(strBuffer, 255)
    
    If strAns <> 0 Then
        GetNameMachine = Left$(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    End If
        
HandleError:
    Call LogSystem("ERROR", "GetNameMachine", Err.Number, Err.Description)
End Function

Public Function GetUserMachine() As String
    On Error GoTo HandleError
    
    Dim UserName As String

    UserName = Environ("USERNAME")

    If UserName <> "" Then
       GetUserMachine = UserName
    End If
        
HandleError:
    Call LogSystem("ERROR", "GetUserMachine", Err.Number, Err.Description)
End Function

Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

Function GetCodeFile() As Long
    On Error GoTo HandleError

    Dim codeFile As Long
    
    gSQLcommand = ""

    gSQLcommand = "INSERT INTO [dbo].[TB_Fichier]"
    gSQLcommand = gSQLcommand & "([NOM_FICHIER]"
    gSQLcommand = gSQLcommand & ",[CHARGE_PAR]"
    gSQLcommand = gSQLcommand & ",[TOTAL_LIGNES]"
    gSQLcommand = gSQLcommand & ",[TOTAL_CLIENTS])"
    gSQLcommand = gSQLcommand & " Values "
    gSQLcommand = gSQLcommand & " ('" & gFileName & "'"
    gSQLcommand = gSQLcommand & " ,'" & gUserMachine & "'"
    gSQLcommand = gSQLcommand & " ," & (gTotalLines - 1) & ""
    gSQLcommand = gSQLcommand & " ," & gQtyCustomerValid & ")"

    gConnectionDB.Open gStringConnection
    gConnectionDB.Execute gSQLcommand
    
    
    gSQLcommand = "SELECT @@IDENTITY AS IDENTITY_"
    
    gRecordsetDB.Open gSQLcommand, gConnectionDB, adOpenStatic, adLockReadOnly
    
    If gRecordsetDB.RecordCount > 0 Then
        codeFile = gRecordsetDB.Fields!IDENTITY_
    Else
        codeFile = 0
    End If
    
    gConnectionDB.Close
    
    GetCodeFile = codeFile

HandleError:
    Call LogSystem("ERROR", "GetCodeFile", Err.Number, Err.Description)
End Function
