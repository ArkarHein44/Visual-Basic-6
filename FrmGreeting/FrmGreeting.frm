VERSION 5.00
Begin VB.Form FrmGreeting 
   Caption         =   "Greeting Program"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdGreet 
      Caption         =   "&Greet Me"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtGreet 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "FrmGreeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    txtGreet.Text = ""
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdGreet_Click()
    txtGreet.Text = "Hello, How are you getting on?"
    
End Sub

