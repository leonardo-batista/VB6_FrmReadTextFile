VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class cFileINI

Private pPathFile           As String
Private pTitle              As String
Private pSystem             As String
Private pCampaign           As String
Private pEmail              As String
Private pHeader             As String
Private pDelimiterColumn    As String
Private pDataSource         As String
Private pDatabase           As String
Private pUser               As String
Private pPassword           As String

Public Property Get PathFile() As String
    PathFile = App.Path & "\config.ini"
End Property

Public Property Get Title() As String
    Title = pTitle
End Property

Public Property Let Title(Value As String)
    pTitle = Value
End Property

Public Property Get System() As String
    System = pSystem
End Property

Public Property Let System(Value As String)
    pSystem = Value
End Property

Public Property Get Campaign() As String
    Campaign = pCampaign
End Property

Public Property Let Campaign(Value As String)
    pCampaign = Value
End Property

Public Property Get Email() As String
    Email = pEmail
End Property

Public Property Let Email(Value As String)
    pEmail = Value
End Property

Public Property Get Header() As String
    Header = pHeader
End Property

Public Property Let Header(Value As String)
    pHeader = Value
End Property

Public Property Get DelimiterColumn() As String
    DelimiterColumn = pDelimiterColumn
End Property

Public Property Let DelimiterColumn(Value As String)
    pDelimiterColumn = Value
End Property

Public Property Get DataSource() As String
    DataSource = pDataSource
End Property

Public Property Let DataSource(Value As String)
    pDataSource = Value
End Property

Public Property Get Database() As String
    Database = pDatabase
End Property

Public Property Let Database(Value As String)
    pDatabase = Value
End Property

Public Property Get User() As String
    User = pUser
End Property

Public Property Let User(Value As String)
    pUser = Value
End Property

Public Property Get Password() As String
    Password = pPassword
End Property

Public Property Let Password(Value As String)
    pPassword = Value
End Property
