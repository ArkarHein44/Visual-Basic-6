VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCommandDialogControl 
   Caption         =   "Using Common Dialog Control"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
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
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
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
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
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
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmCommandDialogControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisplay_Click()
    If Option1(0).Value Then
        CommonDialog1.ShowOpen
        Text1.Text = CommonDialog1.FileName
    ElseIf Option1(1).Value Then
        CommonDialog1.ShowSave
        Text1.Text = CommonDialog1.FileName
    ElseIf Option1(2).Value Then
        CommonDialog1.ShowColor
        Text1.ForeColor = CommonDialog1.Color
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Paint()
    Dim i As Integer
    Static FlagFormPainted As Integer
    If FlagFormPainted <> True Then
        For i = 1 To 2
            Option1(i).Top = Option1(i - 1).Top + 350
            Option1(i).Visible = True
        Next i
        Option1(0).Caption = "Open"
        Option1(1).Caption = "Save"
        Option1(2).Caption = "Show Dlg Color"
        FlagFormPainted = True
    End If
        
End Sub
