Attribute VB_Name = "moduleDBcustomer"
Option Explicit

Public Function FindByIdCustomer(ByVal idCustomer As Long) As cCustomer
    On Error GoTo HandleError

    Dim sqlCommand As String
    gSQLcommand = ""

    Dim customer As cCustomer
    Set customer = New cCustomer


    FindById customer

HandleError:
    Call LogSystem("ERROR", "FindById", Err.Number, Err.Description)
End Function

Public Sub DeleteCustomer(ByVal customer As cCustomer)
    On Error GoTo HandleError

    Dim sqlCommand As String
    gSQLcommand = ""

HandleError:
    Call LogSystem("ERROR", "Customer Delete", Err.Number, Err.Description)
End Sub

Public Sub InsertCustomer(ByVal customer As cCustomer)
    On Error GoTo HandleError

    Dim sqlCommand As String
    gSQLcommand = ""

    sqlCommand = "INSERT INTO dbo.TB_Client"
    sqlCommand = sqlCommand & "(PRENOM"
    sqlCommand = sqlCommand & ",NOM"
    sqlCommand = sqlCommand & ",DATA_NAISSANCE"
    sqlCommand = sqlCommand & ",EMAIL"
    sqlCommand = sqlCommand & ",NAS"
    sqlCommand = sqlCommand & ",TELEPHONE1"
    sqlCommand = sqlCommand & ",TELEPHONE2"
    sqlCommand = sqlCommand & ",CODE_POSTAL"
    sqlCommand = sqlCommand & ",NUMERO"
    sqlCommand = sqlCommand & ",COMPLEMENT"
    sqlCommand = sqlCommand & ",ADRESSE"
    sqlCommand = sqlCommand & ",VILLE"
    sqlCommand = sqlCommand & ",PROVINCE"
    sqlCommand = sqlCommand & ",CREATE_PAR"
    sqlCommand = sqlCommand & ",ID_FICHIER)"
    sqlCommand = sqlCommand & " Values"
    sqlCommand = sqlCommand & " ('" & customer.Prenom & "'"
    sqlCommand = sqlCommand & " ,'" & customer.Nom & "'"
    sqlCommand = sqlCommand & " ,'" & customer.BirthDate & "'"
    sqlCommand = sqlCommand & " ,'" & customer.Email & "'"
    sqlCommand = sqlCommand & " ,'" & customer.Nas & "'"
    sqlCommand = sqlCommand & " ,'" & customer.Telephone1 & "'"
    sqlCommand = sqlCommand & " ,'" & customer.Telephone2 & "'"
    sqlCommand = sqlCommand & " ,'" & customer.CodePostal & "'"
    sqlCommand = sqlCommand & " ,'" & customer.Number & "'"
    sqlCommand = sqlCommand & " ,'" & customer.Complement & "'"
    sqlCommand = sqlCommand & " ,'" & customer.Address & "'"
    sqlCommand = sqlCommand & " ,'" & customer.City & "'"
    sqlCommand = sqlCommand & " ,'" & customer.UnitFed & "'"
    sqlCommand = sqlCommand & " ,'" & gUserMachine & "'"
    sqlCommand = sqlCommand & " ," & customer.IdFile & ")"

    gConnectionDB.Open gStringConnection
    gConnectionDB.Execute sqlCommand
    sqlCommand = ""
    
HandleError:
    gConnectionDB.Close
    
    If Err.Number = 0 Then
        gQtyInsertedDB = gQtyInsertedDB + 1
    End If
    
    Call LogSystem("ERROR", "Customer Insert: " & sqlCommand, Err.Number, Err.Description)
End Sub

Public Sub QtyCustomerInserted()
    On Error GoTo HandleError

    Dim sqlCommand As String
    gSQLcommand = ""

    gSQLcommand = "UPDATE [dbo].[TB_Fichier] SET [DT_MISE_A_JOUR] = GETDATE(), [TOTAL_CLIENTS] = " & gQtyInsertedDB & " WHERE ID_FICHIER = " & CInt(gFileCode)

    gConnectionDB.Open gStringConnection
    gConnectionDB.Execute gSQLcommand

HandleError:
    gConnectionDB.Close
    Call LogSystem("ERROR", "File Update: " & sqlCommand, Err.Number, Err.Description)
End Sub
