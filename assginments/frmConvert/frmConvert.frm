VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Convert to Uppercase"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   953
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    txtUser.Text = ""
    txtUser.SetFocus
End Sub

Private Sub cmdExit_Click()
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Are you sure you want to exit", vbExclamation + vbYesNo, "Confirm Exit")
    If confirm = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
    txtUser.MaxLength = 15
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If Len(txtUser.Text) >= 15 And KeyAscii <> vbKeyBack Then
        MsgBox "Maximun length of 15 characters reached!", vbExclamation, "Input Limit"
        KeyAscii = 0
        Exit Sub
    End If
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

