VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Read Text File"
   ClientHeight    =   6210
   ClientLeft      =   7230
   ClientTop       =   4530
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmDialog1 
      Left            =   120
      Top             =   120
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

Private Sub Exit_Click()
    If MsgBox("Are you sure you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
        End
    End If
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
    OpenFileTxt
End Sub

Private Sub OpenFileTxt()

    cmDialog1.Filter = ".txt File (*.txt)|*.txt"
    cmDialog1.DefaultExt = "txt"
    cmDialog1.DialogTitle = "Select your file"
    cmDialog1.CancelError = True
    cmDialog1.ShowOpen

End Sub
