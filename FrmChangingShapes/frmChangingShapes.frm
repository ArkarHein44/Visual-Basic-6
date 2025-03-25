VERSION 5.00
Begin VB.Form frmShape 
   Caption         =   "Changing Shapes"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Shape"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1313
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaxLength       =   1
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Type a number between 0 to 5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
    Shape1.Shape = Val(Text1.Text)
End Sub

Private Sub Text1_Change()
    If Not IsNumeric(Text1.Text) Then
        Text1.Text = ""
    ElseIf Val(Text1.Text) < 0 Then
        Text1.Text = 0
    ElseIf Val(Text1.Text) > 5 Then
        Text1.Text = 5
    End If
End Sub
