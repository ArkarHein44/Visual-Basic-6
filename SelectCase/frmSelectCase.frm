VERSION 5.00
Begin VB.Form frmSelectCase 
   Caption         =   "Select Case"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
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
      Left            =   2543
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
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
      Left            =   743
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtIn 
      Alignment       =   2  'Center
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
      Left            =   803
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frmSelectCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdShow_Click()
    Dim X As String
    X = Trim(txtIn.Text)
    
    Select Case X
        Case "One"
        MsgBox "Your input is 1."
        
        Case "Two"
        MsgBox "Your input is 2."
        
        Case Else
        MsgBox "Your input is neither 1 nor 2."
    End Select
End Sub

