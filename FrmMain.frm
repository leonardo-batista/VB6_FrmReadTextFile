VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
      Height          =   6855
      Left            =   6120
      TabIndex        =   20
      Top             =   240
      Width           =   5535
      Begin VB.CommandButton btnStartProcess 
         Caption         =   "Start Process"
         Height          =   735
         Left            =   3120
         TabIndex        =   23
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optValidatonDatabase 
         Caption         =   "Validation + Load Database"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton optValidation 
         Caption         =   "Validation, only"
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1455
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
      Top             =   1320
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
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   3000
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
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   240
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
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblWorkstation 
         Caption         =   "Workstation"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
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

    If cmDialog1.FileName <> "" Then
    
        If optValidation.Value = False And optValidatonDatabase.Value = False Then
            MsgBox "Please, select one Process Option !!!", vbExclamation, "Alert - Process"
        Else
            'Function or Methode Here
        End If
        
    Else
        MsgBox "Please, select one file !!!", vbExclamation, "Alert - File"
    End If

End Sub

Private Sub Exit_Click()
    If MsgBox("Are you sure you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
    LoadConfigINI
    GetInformation
    FormInformation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
        Unload Me
    Else
        Cancel = True
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
    
HandleError:
    Debug.Print Err.Number & " " & Err.Description
    
End Sub

Private Sub FormInformation()

    txtDatabaseConnection.Text = gDateAccess
    txtWorkstation.Text = gNameMachine
    txtUser.Text = gUserMachine
End Sub

Private Sub CleanFieldsFile()
    txtPathFile.Text = ""
    txtFileName.Text = ""
    txtHeaderFile.Text = ""
    txtColumnDelimiter.Text = ""
    txtTotalCustomers.Text = ""
    txtTotalLines.Text = ""
    optValidation.Value = False
    optValidatonDatabase.Value = False
End Sub
