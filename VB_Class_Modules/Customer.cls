VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class cCustomer

Private pPrenom      As String
Private pNom         As String
Private pBirthDate   As String
Private pEmail       As String
Private pNas         As String
Private pTelephone1  As String
Private pTelephone2  As String
Private pCodePostal  As String
Private pNumber      As String
Private pComplement  As String
Private pAddress     As String
Private pCity        As String
Private pUnitFed     As String

Public Property Get Prenom() As String
    Prenom = pPrenom
End Property

Public Property Let Prenom(Value As String)
    pPrenom = Value
End Property

Public Property Get Nom() As String
    Nom = pNom
End Property

Public Property Let Nom(Value As String)
    pNom = Value
End Property

Public Property Get BirthDate() As String
    BirthDate = pBirthDate
End Property

Public Property Let BirthDate(Value As String)
    pBirthDate = Value
End Property

Public Property Get Email() As String
    Email = pEmail
End Property

Public Property Let Email(Value As String)
    pEmail = Value
End Property

Public Property Get Nas() As String
    Nas = pNas
End Property

Public Property Let Nas(Value As String)
    pNas = Value
End Property

Public Property Get Telephone1() As String
    Telephone1 = pTelephone1
End Property

Public Property Let Telephone1(Value As String)
    pTelephone1 = Value
End Property

Public Property Get Telephone2() As String
    Telephone2 = pTelephone2
End Property

Public Property Let Telephone2(Value As String)
    pTelephone2 = Value
End Property

Public Property Get CodePostal() As String
    CodePostal = pCodePostal
End Property

Public Property Let CodePostal(Value As String)
    pCodePostal = Value
End Property

Public Property Get Number() As String
    Number = pNumber
End Property

Public Property Let Number(Value As String)
    pNumber = Value
End Property

Public Property Get Complement() As String
    Complement = pComplement
End Property

Public Property Let Complement(Value As String)
    pComplement = Value
End Property

Public Property Get Address() As String
    Address = pAddress
End Property

Public Property Let Address(Value As String)
    pAddress = Value
End Property

Public Property Get City() As String
    City = pCity
End Property

Public Property Let City(Value As String)
    pCity = Value
End Property

Public Property Get UnitFed() As String
    UnitFed = pUnitFed
End Property

Public Property Let UnitFed(Value As String)
    pUnitFed = Value
End Property
