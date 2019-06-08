Attribute VB_Name = "moduleCustomer"
Option Explicit

'--------------------------------------------------
'FUNCTIONS / SUBROUTINES FOR CUSTOMER (VALIDATION)
'--------------------------------------------------

Public Sub CustomerValidation(pCustomer As cCustomer)
    On Error GoTo HandleError
   
    Dim isValidBirthDate As Boolean
    Dim isValidTelephone1 As Boolean
    Dim isValidTelephone2 As Boolean
   
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
    If BirthDate(pCustomer.BirthDate) Then
        isValidBirthDate = True
        qQtyBirthDateValid = qQtyBirthDateValid + 1
    Else
        isValidBirthDate = False
        qQtyBirthDateInvalid = qQtyBirthDateInvalid + 1
    End If
       
    'Telephone 1
    If Telephone(pCustomer.Telephone1) Then
        isValidTelephone1 = True
        gQtyTelephone1Valid = gQtyTelephone1Valid + 1
    Else
        isValidTelephone1 = False
        gQtyTelephone1Invalid = gQtyTelephone1Invalid + 1
    End If
    
    'Telephone 2
    If Telephone(pCustomer.Telephone2) Then
        isValidTelephone2 = True
        gQtyTelephone2Valid = gQtyTelephone2Valid + 1
    Else
        isValidTelephone2 = False
        gQtyTelephone2Invalid = gQtyTelephone2Invalid + 1
    End If
    
    'E-mail
    If Email(pCustomer.Email) Then
        gQtyEmailValid = gQtyEmailValid + 1
    Else
        gMessageValidation = gMessageValidation & "|E-mail Invalid"
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
      
    'Customer valid: +18 and 1 Phone Number (minimum)
    If IsAdult(pCustomer.BirthDate) = True And (isValidTelephone1 = True Or isValidTelephone2 = True) Then
        gQtyCustomerValid = gQtyCustomerValid + 1
        pCustomer.IsValid = True
    Else
        gQtyCustomerInvalid = gQtyCustomerInvalid + 1
        pCustomer.IsValid = False
    End If
    
    validationCustomer.LineFile = pCustomer.LineFile
    validationCustomer.Prenom = pCustomer.Prenom
    validationCustomer.Nom = pCustomer.Nom
    validationCustomer.Nas = pCustomer.Nas
    validationCustomer.IsValid = pCustomer.IsValid
    validationCustomer.MsgValidation = gMessageValidation
    
    gValidationCustomers.Add Item:=validationCustomer
       
HandleError:
    Call LogSystem("ERROR", "ValidationCustomerList", Err.Number, Err.Description)
End Sub

Private Function Nom(pNom As String) As Boolean

    Dim IsValid As Boolean
    IsValid = True
    
    If Trim(pNom) = "" Then
        IsValid = False
    End If

    If IsValid Then
        Nom = IsValid
    Else
        Nom = IsValid
    End If
      
End Function

Private Function Prenom(pPrenom As String) As Boolean

    Dim IsValid As Boolean
    IsValid = True
    
    If Trim(pPrenom) = "" Then
        IsValid = False
    End If

    If IsValid Then
        Prenom = IsValid
    Else
        Prenom = IsValid
    End If
    
End Function

Private Function Nas(pNas As String) As Boolean

    Dim IsValid As Boolean
    IsValid = True
    
    If Trim(pNas) = "" Then
        IsValid = False
    End If

    If Len(Trim(pNas)) <> 9 Then
        IsValid = False
    End If

    If IsValid Then
        Nas = IsValid
    Else
        Nas = IsValid
    End If

End Function

Private Function BirthDate(pBirthDate As String) As Boolean

    Dim IsValid As Boolean
    IsValid = True
    
    If Trim(pBirthDate) = "" Then
        IsValid = False
    End If

    If Len(Trim(pBirthDate)) <> 10 Then
        IsValid = False
    End If

    'If Not Format(pBirthDate, "yyyy-MM-dd") Then
    '    IsValid = False
    'End If

    If Not IsDate(Trim(pBirthDate)) Then
        IsValid = False
    End If

    If IsValid Then
        BirthDate = IsValid
    Else
        BirthDate = IsValid
    End If

End Function

Private Function IsAdult(pBirthDate As String) As Boolean

    Dim IsValid As Boolean
    IsValid = True

    Dim dtCustomer As Date
    Dim age As Long
    
    If pBirthDate = "" Or Len(pBirthDate) <> 10 Then
        IsAdult = False
        Exit Function
    End If
    
    dtCustomer = CDate(Format(pBirthDate, "yyyy-MM-dd"))
    age = DateDiff("yyyy", dtCustomer, Now)

    If age < 18 Then
        IsAdult = False
    End If

    If IsValid Then
        IsAdult = IsValid
    Else
        IsAdult = IsValid
    End If

End Function

Private Function Telephone(pTelephone As String) As Boolean

    Dim phone As String
    Dim IsValid As Boolean
    IsValid = True
    
    phone = Trim(pTelephone)
       
    'If IsNumeric(CLng(phone)) Then
    '    IsValid = False
    'End If
    
    If Len(phone) <> 10 Then
        IsValid = False
    End If
    
    If IsValid Then
        Telephone = IsValid
    Else
        Telephone = IsValid
    End If
    
End Function

Private Function Email(pEmail As String) As Boolean

    Dim IsValid As Boolean
    Dim strDomainType As String
    Dim strDomainName As String
    Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
    Dim i As Integer
    
    IsValid = Not InStr(1, pEmail, Chr(34)) > 0 'Check to see if there is a double quote
    If Not IsValid Then GoTo ExitFunction
    
    IsValid = Not InStr(1, pEmail, "..") > 0 'Check to see if there are consecutive dots
    If Not IsValid Then GoTo ExitFunction
    
    ' Check for invalid characters.
    If Len(pEmail) > Len(sInvalidChars) Then
        For i = 1 To Len(sInvalidChars)
            If InStr(pEmail, Mid(sInvalidChars, i, 1)) > 0 Then
                IsValid = False
                GoTo ExitFunction
            End If
        Next
    Else
        For i = 1 To Len(pEmail)
            If InStr(sInvalidChars, Mid(pEmail, i, 1)) > 0 Then
                IsValid = False
                GoTo ExitFunction
            End If
        Next
    End If
    
    If InStr(1, pEmail, "@") > 1 Then 'Check for an @ symbol
        IsValid = Len(Left(pEmail, InStr(1, pEmail, "@") - 1)) > 0
    Else
        IsValid = False
    End If
    If Not IsValid Then GoTo ExitFunction
    
    pEmail = Right(pEmail, Len(pEmail) - InStr(1, pEmail, "@"))
    IsValid = Not InStr(1, pEmail, "@") > 0 'Check to see if there are too many @'s
    If Not IsValid Then GoTo ExitFunction
    
    strDomainType = Right(pEmail, Len(pEmail) - InStr(1, pEmail, "."))
    IsValid = Len(strDomainType) > 0 And InStr(1, pEmail, ".") < Len(pEmail)
    If Not IsValid Then GoTo ExitFunction
    
    pEmail = Left(pEmail, Len(pEmail) - Len(strDomainType) - 1)
    Do Until InStr(1, pEmail, ".") <= 1
        If Len(pEmail) >= InStr(1, pEmail, ".") Then
            pEmail = Left(pEmail, Len(pEmail) - (InStr(1, pEmail, ".") - 1))
        Else
            IsValid = False
            GoTo ExitFunction
        End If
    Loop
    If pEmail = "." Or Len(pEmail) = 0 Then IsValid = False
    
ExitFunction:
    Email = IsValid

End Function

Private Function FedUnit(pFedUnit As String) As Boolean

    Dim IsValid As Boolean
    IsValid = True

    If Trim(pFedUnit) = "" Then
        IsValid = False
    End If

    If Len(Trim(pFedUnit)) > 2 Then
        IsValid = False
    End If

    Select Case UCase(pFedUnit)
        Case Is = "AB"
            '
        Case Is = "BC"
            '
        Case Is = "PE"
            '
        Case Is = "MB"
            '
        Case Is = "NB"
            '
        Case Is = "NS"
            '
        Case Is = "NU"
            '
        Case Is = "ON"
            '
        Case Is = "QC"
            '
        Case Is = "SK"
            '
        Case Is = "NL"
            '
        Case Is = "NT"
            '
        Case Is = "YT"
            '
        Case Else
            IsValid = False
      End Select
  
    If IsValid Then
        FedUnit = IsValid
    Else
        FedUnit = IsValid
    End If

End Function

Private Function PostalCode(pPostalCode As String) As Boolean

    Dim IsValid As Boolean
    IsValid = True
    
    If Trim(pPostalCode) = "" Then
        IsValid = False
    End If

    If IsValid Then
        PostalCode = IsValid
    Else
        PostalCode = IsValid
    End If

End Function
