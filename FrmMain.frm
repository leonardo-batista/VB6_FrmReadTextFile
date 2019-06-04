VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Read Text File"
   ClientHeight    =   7290
   ClientLeft      =   7230
   ClientTop       =   4530
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameProcessFile 
      Caption         =   "Process File..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   6120
      TabIndex        =   20
      Top             =   120
      Width           =   5535
      Begin MSComctlLib.ProgressBar pgrbProcessFile 
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.CommandButton btnStartProcess 
         Caption         =   "Start Process"
         Height          =   735
         Left            =   3120
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optValidatonDatabase 
         Caption         =   "Validation + Load Database"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton optValidation 
         Caption         =   "Validation, only"
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label30 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4560
         TabIndex        =   68
         Top             =   6720
         Width           =   660
      End
      Begin VB.Label Label29 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4560
         TabIndex        =   67
         Top             =   6360
         Width           =   660
      End
      Begin VB.Label Label28 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4560
         TabIndex        =   66
         Top             =   6000
         Width           =   660
      End
      Begin VB.Label Label27 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4560
         TabIndex        =   65
         Top             =   5640
         Width           =   660
      End
      Begin VB.Label Label26 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   64
         Top             =   6720
         Width           =   660
      End
      Begin VB.Label Label25 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   63
         Top             =   6360
         Width           =   660
      End
      Begin VB.Label Label24 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   62
         Top             =   6000
         Width           =   660
      End
      Begin VB.Label Label23 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   61
         Top             =   5640
         Width           =   660
      End
      Begin VB.Label Label22 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   60
         Top             =   5280
         Width           =   660
      End
      Begin VB.Label Label21 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   59
         Top             =   4440
         Width           =   660
      End
      Begin VB.Label Label20 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   58
         Top             =   4080
         Width           =   660
      End
      Begin VB.Label Label19 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   57
         Top             =   3720
         Width           =   660
      End
      Begin VB.Label Label18 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   56
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label Label17 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   55
         Top             =   4440
         Width           =   660
      End
      Begin VB.Label Label16 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   54
         Top             =   4080
         Width           =   660
      End
      Begin VB.Label Label15 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   53
         Top             =   3720
         Width           =   660
      End
      Begin VB.Label Label14 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   52
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label Label13 
         Caption         =   "000000"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   51
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label Label12 
         Caption         =   "?"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4560
         TabIndex        =   50
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label11 
         Caption         =   "?"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   49
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label5 
         Caption         =   "Fed. Unit .................."
         Height          =   255
         Index           =   13
         Left            =   2880
         TabIndex        =   48
         Top             =   6720
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Telephone 2 ............"
         Height          =   255
         Index           =   12
         Left            =   2880
         TabIndex        =   47
         Top             =   6360
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Birth Date ................"
         Height          =   255
         Index           =   11
         Left            =   2880
         TabIndex        =   46
         Top             =   6000
         Width           =   1500
      End
      Begin VB.Label Label10 
         Caption         =   "Prenom ...................."
         Height          =   255
         Left            =   2880
         TabIndex        =   45
         Top             =   5640
         Width           =   1500
      End
      Begin VB.Label Label9 
         Caption         =   "Customer ................."
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   5280
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Code Postal ............."
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   43
         Top             =   6720
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Telephone 1 ............"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   42
         Top             =   6360
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "NAS ........................."
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   41
         Top             =   6000
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Nom ........................."
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   40
         Top             =   5640
         Width           =   1500
      End
      Begin VB.Label Label8 
         Caption         =   "Report....... .............."
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "NAS ........................."
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   38
         Top             =   3720
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Birth Date ................"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   36
         Top             =   3720
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Fed. Unit .................."
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   35
         Top             =   4440
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Code Postal ............."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   34
         Top             =   4440
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Telephone 2 ............"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   33
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Telephone 1 ............"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "Customer ................."
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label Label6 
         Caption         =   "Prenom ...................."
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Nom ........................."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Column Delimiter ....."
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Header File .............."
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "Invalid ........................................................................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   4920
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Valid ............................................................................"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2640
         Width           =   5175
      End
      Begin VB.Label lblFileResume 
         Caption         =   "File Resume ................................................................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   5175
      End
   End
   Begin VB.TextBox txtPathFile 
      Height          =   380
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "File"
      Height          =   380
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtWorkstation 
      Height          =   380
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtUser 
      Height          =   380
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Frame FrameSelectFile 
      Caption         =   "Select your file (txt)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   5775
      Begin VB.TextBox txtHeaderFile 
         Height          =   380
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtColumnDelimiter 
         Height          =   380
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtTotalCustomers 
         Height          =   380
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox txtTotalLines 
         Height          =   380
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2880
         Width           =   3255
      End
      Begin VB.TextBox txtFileName 
         Height          =   380
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblHeaderFile 
         Caption         =   "Header File"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblColumnDelimiter 
         Caption         =   "Column Delimiter"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblTotalCustomers 
         Caption         =   "Total Customers"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label lblTotalLines 
         Caption         =   "Total Lines"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblFileName 
         Caption         =   "File"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Frame frameSystemInfo 
      Caption         =   "System Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtDatabaseConnection 
         Height          =   380
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblUser 
         Caption         =   "User"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblWorkstation 
         Caption         =   "Workstation"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblDatabaseConnection 
         Caption         =   "Database Connection Test"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog cmDialog1 
      Left            =   11280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFile_Click()
    CleanFieldsFile
    OpenFileTxt
End Sub

Private Sub btnStartProcess_Click()
    On Error GoTo HandleError
    
    If cmDialog1.FileName <> "" Then
    
        If optValidation.Value = False And optValidatonDatabase.Value = False Then
            MsgBox "Please, select one Process Option !!!", vbExclamation, "Alert - Process"
        Else
            If optValidation.Value = True Then
                'Function or Methode Here
                ConvertFileToCustomer
                ValidationCustomerList
            End If
            
            If optValidatonDatabase.Value = True Then
                'Function or Methode Here
            End If
            
            Dim count As Integer
            
            For count = 1 To gTotalLines
                        
                pgrbProcessFile.Value = pgrbProcessFile.Value + 1
            
            Next
            
        End If
        
    Else
        MsgBox "Please, select one file !!!", vbExclamation, "Alert - File"
    End If

HandleError:
    Call LogSystem("ERROR", "btnStartProcess_Click", Err.Number, Err.Description)
End Sub

Private Sub Exit_Click()
    If MsgBox("Are you sure you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
        LogSystem "INFO", "Exit_Click", 0, "Exit of System"
        End
    End If
End Sub

Private Sub Form_Load()
    SystemDirectory
    LoadConfigINI
    GetInformation
    FormInformation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
        Unload Me
    Else
        Cancel = True
        LogSystem "INFO", "Form_QueryUnload", 0, "Exit of System"
        Exit Sub
    End If
End Sub

Private Sub Open_Click()
    CleanFieldsFile
    OpenFileTxt
End Sub

Private Sub OpenFileTxt()
    On Error GoTo HandleError
    
    cmDialog1.Filter = ".txt File (*.txt)|*.txt"
    cmDialog1.DefaultExt = "txt"
    cmDialog1.DialogTitle = "Select your file"
    cmDialog1.CancelError = True
    cmDialog1.ShowOpen
    
    If cmDialog1.CancelError = False Then
        CleanFieldsFile
    End If
    
    If cmDialog1.FileName <> "" Then
        txtPathFile.Text = cmDialog1.FileName
        txtFileName.Text = GetFileNameFromPath(txtPathFile.Text)
    End If
       
    iFile = FreeFile
        
    Dim sLineText$
    Dim customer As cCustomer
    
    Open cmDialog1.FileName For Input As #iFile
        
        Do While Not EOF(1)
        
            Input #iFile, sLineText$
            
            sLineText$ = Trim(sLineText$)
            
            If sLineText$ <> "" Then
            
                If gFileHeader = "" Then
                    gFileHeader = sLineText$
                End If
                
                gTotalLines = gTotalLines + 1
                
            End If
                        
            DoEvents
            
        Loop
        
    Close #iFile
    
    txtTotalLines.Text = gTotalLines
    txtTotalCustomers.Text = gTotalLines - 1
    txtColumnDelimiter.Text = gFileIni.DelimiterColumn
    txtHeaderFile.Text = gFileHeader
    pgrbProcessFile.Max = gTotalLines
    
HandleError:
    Call LogSystem("ERROR", "OpenFileTxt", Err.Number, Err.Description)
End Sub

Private Sub FormInformation()

    txtDatabaseConnection.Text = gDateAccess
    txtWorkstation.Text = gNameMachine
    txtUser.Text = gUserMachine
End Sub

Private Sub CleanFieldsFile()
    pgrbProcessFile.Value = 0
    txtPathFile.Text = ""
    txtFileName.Text = ""
    txtHeaderFile.Text = ""
    txtColumnDelimiter.Text = ""
    txtTotalCustomers.Text = ""
    txtTotalLines.Text = ""
    optValidation.Value = False
    optValidatonDatabase.Value = False
End Sub

Private Sub ConvertFileToCustomer()

iFile = FreeFile
        
    Dim sLineText$
    Dim customer As cCustomer
    
    Open cmDialog1.FileName For Input As #iFile
        
        Do While Not EOF(1)
        
            Input #iFile, sLineText$
            
            If sLineText$ <> "" Then
            
                If gFileName = "" Then
                    gFileName = sLineText$
                End If
                
                lineValue = Split(sLineText$, gFileIni.DelimiterColumn)
                
                Set customer = New cCustomer
    
                customer.Prenom = lineValue(0)
                customer.Nom = lineValue(1)
                customer.BirthDate = lineValue(2)
                customer.Email = lineValue(3)
                customer.Nas = lineValue(4)
                customer.Telephone1 = lineValue(5)
                customer.Telephone2 = lineValue(6)
                customer.CodePostal = lineValue(7)
                customer.Number = lineValue(8)
                customer.Complement = lineValue(9)
                customer.Address = lineValue(10)
                customer.City = lineValue(11)
                customer.UnitFed = lineValue(12)
                
                gResultCustomers.Add Item:=customer
                
                End If
          
            DoEvents
            
        Loop
        
    Close #iFile

End Sub

Private Sub ValidationCustomerList()
    On Error GoTo HandleError
    
    Dim count As Integer
    Dim itemCustomer As cCustomer
    
    Set itemCustomer = Nothing
    
    For Each itemCustomer In gResultCustomers
        Debug.Print itemCustomer.Nom
    Next itemCustomer
   
HandleError:
    Call LogSystem("ERROR", "ValidationCustomerList", Err.Number, Err.Description)
End Sub
