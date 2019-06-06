Attribute VB_Name = "moduleCustomer"
Option Explicit

'--------------------------------------------------
'FUNCTIONS / SUBROUTINES FOR CUSTOMER (VALIDATION)
'--------------------------------------------------

Public Sub CustomerValidation(pCustomer As cCustomer)
    On Error GoTo HandleError
   
    gMessageValidation = "|"
   
    Dim validationCustomer As cValidation
    Set validationCustomer = New cValidation
    
    'Nom
    If Nom(pCustomer.Nom) Then
        gQtyNomValid = gQtyNomValid + 1
    Else
        gQtyNomInvalid = gQtyNomInvalid + 1
    End If
    
    'Prenom
    If Prenom(pCustomer.Prenom) Then
        gQtyPrenomValid = gQtyPrenomValid + 1
    Else
        gQtyPrenomInvalid = gQtyPrenomInvalid + 1
    End If
    
    'NAS
    If Nas(pCustomer.Nas) Then
        gQtyNASValid = gQtyNASValid + 1
    Else
        gQtyNASInvalid = gQtyNASInvalid + 1
    End If
    
    'Birth Date
    If Nas(pCustomer.BirthDate) Then
        'gQtyBir = gQtyNASValid + 1
    Else
        'gQtyNASInvalid = gQtyNASInvalid + 1
    End If
    
    'Telephone 1
    If Telephone(pCustomer.Telephone1) Then
        gQtyTelephone1Valid = gQtyTelephone1Valid + 1
    Else
        gQtyTelephone1Invalid = gQtyTelephone1Invalid + 1
    End If
    
    'Telephone 2
    If Telephone(pCustomer.Telephone2) Then
        gQtyTelephone2Valid = gQtyTelephone2Valid + 1
    Else
        gQtyTelephone2Invalid = gQtyTelephone2Invalid + 1
    End If
    
    'E-mail
    If Email(pCustomer.Email) Then
        gQtyEmailValid = gQtyEmailValid + 1
    Else
        gQtyEmailInvalid = gQtyEmailInvalid + 1
    End If
    
    'Fed Unit
    If FedUnit(pCustomer.UnitFed) Then
        gQtyFedUnitValid = gQtyFedUnitValid + 1
    Else
        gQtyFedUnitInvalid = gQtyFedUnitInvalid + 1
    End If
    
    If PostalCode(pCustomer.CodePostal) Then
        gQtyPostalCodeValid = gQtyPostalCodeValid + 1
    Else
        gQtyPostalCodeInvalid = gQtyPostalCodeInvalid + 1
    End If
        
    validationCustomer.LineFile = pCustomer.LineFile
    validationCustomer.Prenom = pCustomer.Prenom
    validationCustomer.Nom = pCustomer.Nom
    validationCustomer.Nas = pCustomer.Nas
    'validationCustomer.IsValid = False
    validationCustomer.MsgValidation = gMessageValidation
    
    gValidationCustomers.Add Item:=validationCustomer
    
    
    
    pCustomer.isValid = True

    'Debug.Print pCustomer.Nas
    'gValidationCustomers
       
HandleError:
    Call LogSystem("ERROR", "ValidationCustomerList", Err.Number, Err.Description)
End Sub

Private Function Nom(pNom As String) As Boolean

    Dim isValid As Boolean
    isValid = True
    
    If Trim(pNom) = "" Then
        'gMessageValidation = gMessageValidation & "|Nom is Empty"
        isValid = False
    End If

    If isValid Then
        Nom = isValid
    Else
        Nom = isValid
    End If
      
End Function

Private Function Prenom(pPrenom As String) As Boolean

    Dim isValid As Boolean
    isValid = True
    
    If Trim(pPrenom) = "" Then
        isValid = False
        'gMessageValidation = gMessageValidation & "|Nom is Empty"
    End If

    If isValid Then
        Prenom = isValid
    Else
        Prenom = isValid
    End If
    
End Function

Private Function Nas(pNas As String) As Boolean

    Dim isValid As Boolean
    isValid = True
    
    If Trim(pNas) = "" Then
        isValid = False
        'gMessageValidation = gMessageValidation & "|Nom is Empty"
    End If

    If isValid Then
        Nas = isValid
    Else
        Nas = isValid
    End If

End Function

Private Function BirthDate(pBirthDate As String) As Boolean

    Dim isValid As Boolean
    isValid = True
    
    If Trim(pBirthDate) = "" Then
        isValid = False
        'gMessageValidation = gMessageValidation & "|Nom is Empty"
    End If

    If isValid Then
        BirthDate = isValid
    Else
        BirthDate = isValid
    End If

End Function

Private Function Telephone(pTelephone As String) As Boolean

    Dim isValid As Boolean
    isValid = True
    
    If Trim(pTelephone) = "" Then
        isValid = False
        'gMessageValidation = gMessageValidation & "|Nom is Empty"
    End If

    If isValid Then
        Telephone = isValid
    Else
        Telephone = isValid
    End If
    
End Function

Private Function Email(pEmail As String) As Boolean

    Dim isValid As Boolean
    isValid = True

    If Trim(pEmail) = "" Then
        isValid = False
        'gMessageValidation = gMessageValidation & "|Nom is Empty"
    End If

    If isValid Then
        Email = isValid
    Else
        Email = isValid
    End If

End Function

Private Function FedUnit(pFedUnit As String) As Boolean

    Dim isValid As Boolean
    isValid = True

    If Trim(pFedUnit) = "" Then
        isValid = False
        'gMessageValidation = gMessageValidation & "|Nom is Empty"
    End If

    If isValid Then
        FedUnit = isValid
    Else
        FedUnit = isValid
    End If

End Function

Private Function PostalCode(pPostalCode As String) As Boolean

    Dim isValid As Boolean
    isValid = True
    
    If Trim(pPostalCode) = "" Then
        isValid = False
        'gMessageValidation = gMessageValidation & "|Nom is Empty"
    End If

    If isValid Then
        PostalCode = isValid
    Else
        PostalCode = isValid
    End If

End Function
