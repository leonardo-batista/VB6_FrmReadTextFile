Attribute VB_Name = "moduleCustomer"
Option Explicit


'--------------------------------------------------
'FUNCTIONS / SUBROUTINES FOR CUSTOMER (VALIDATION)
'--------------------------------------------------

Public Function Nom(pNom As String) As Boolean

    If Trim(pNom) = "" Then
        gQtyNomInvalid = gQtyNomInvalid + 1
        Exit Function
    End If

    gQtyNomValid = gQtyNomValid + 1
    Nom = True
    
End Function

Public Function Prenom(pPrenom As String) As Boolean

    If Trim(pPrenom) = "" Then
        Exit Function
    End If

    Prenom = True
    
End Function

Public Function Nas(pNas As String) As Boolean

    If Trim(pNas) = "" Then
    End If

End Function

Public Function BirthDate(pBirthDate As String) As Boolean

    If Trim(pBirthDate) = "" Then
        Exit Function
    End If

    BirthDate = True

End Function

Public Function Telephone(pTelephone As String) As Boolean

    If Trim(pTelephone) = "" Then
        Exit Function
    End If

    Telephone = True

End Function

Public Function Email(pEmail As String) As Boolean

    If Trim(pEmail) = "" Then
        Exit Function
    End If

    Email = True

End Function

Public Function FedUnit(pFedUnit As String) As Boolean

    If Trim(pFedUnit) = "" Then
        Exit Function
    End If

    Province = True

End Function

Public Function PostalCode(pPostalCode As String) As Boolean

    If Trim(pPostalCode) = "" Then
        Exit Function
    End If

    PostalCode = True

End Function
