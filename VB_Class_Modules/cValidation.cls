VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class cValidation

Private pLineFile      As Integer
Private pPrenom        As String
Private pNom           As String
Private pNas           As String
Private pMsgValidation As String

Public Property Get LineFile() As String
    LineFile = pLineFile
End Property

Public Property Let LineFile(Value As String)
    pLineFile = Value
End Property

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

Public Property Get Nas() As String
    Nas = pNas
End Property

Public Property Let Nas(Value As String)
    pNas = Value
End Property

Public Property Get MsgValidation() As String
    MsgValidation = pMsgValidation
End Property

Public Property Let MsgValidation(Value As String)
    pMsgValidation = Value
End Property
