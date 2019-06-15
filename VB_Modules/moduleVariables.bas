Attribute VB_Name = "moduleVariables"
Option Explicit

'--------------------------------------------------
'COMPUTER NAME
'--------------------------------------------------
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'--------------------------------------------------
'USER NAME
'--------------------------------------------------
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'--------------------------------------------------
'DIRECTORY
'--------------------------------------------------
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

'--------------------------------------------------
'API Function to read information from INI File
'--------------------------------------------------
Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long
    
'--------------------------------------------------
'API Function to write information to the INI File
'--------------------------------------------------
Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long
    
Global gFileIni As New cFileIni

'--------------------------------------------------
'Connection Database (ODBC) / Driver SQL
'--------------------------------------------------
Global gConnectionDB As New ADODB.Connection
Global gRecordsetDB As New ADODB.Recordset
Global gSQLcommand As String

'--------------------------------------------------
'VARIABLES TO VALIDATION AND REPORT ABOUT FILE TXT
'--------------------------------------------------
Global gFileCode As Long
Global gFileName As String
Global gFileHeader As String
Global gResultCustomers As New Collection
Global gValidationCustomers As New Collection
Global gLineTotalCustomers As Integer
Global gLinesWithSuccess As Integer
Global gLinesWithError As Integer
Global gTotalLines As Integer

Global gQtyCustomerValid As Integer
Global gQtyNomValid As Integer
Global gQtyPrenomValid As Integer
Global gQtyNASValid As Integer
Global gQtyEmailValid As Integer
Global gQtyTelephone1Valid As Integer
Global gQtyTelephone2Valid As Integer
Global qQtyBirthDateValid As Integer
Global gQtyPostalCodeValid As Integer
Global gQtyFedUnitValid As Integer

Global gQtyCustomerInvalid As Integer
Global gQtyNomInvalid As Integer
Global gQtyPrenomInvalid As Integer
Global gQtyNASInvalid As Integer
Global gQtyEmailInvalid As Integer
Global gQtyTelephone1Invalid As Integer
Global gQtyTelephone2Invalid As Integer
Global qQtyBirthDateInvalid As Integer
Global gQtyPostalCodeInvalid As Integer
Global gQtyFedUnitInvalid As Integer

Global gRegisterIsValid As Boolean
Global gMessageValidation As String

'--------------------------------------------------
'VARIABLES TO LOG
'--------------------------------------------------
Global gUserMachine As String
Global gNameMachine As String
Global gDateAccess As String
