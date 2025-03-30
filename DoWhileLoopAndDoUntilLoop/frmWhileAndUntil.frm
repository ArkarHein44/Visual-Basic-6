VERSION 5.00
Begin VB.Form frmWhileandUntil 
   Caption         =   "Do While Loop And Do Until Loop"
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
      Height          =   495
      Left            =   953
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdUntil 
      Caption         =   "Do UntilLoop"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2573
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdWhileLoop 
      Caption         =   "Do While Loop"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   413
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmWhileandUntil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdUntil_Click()
    n = 1
    Do Until n > 10
        MsgBox n
        n = n + 1
    Loop
End Sub

Private Sub cmdWhileLoop_Click()
    n = 1
    Do While n <= 10
        MsgBox n
        n = n + 1
    Loop
End Sub
