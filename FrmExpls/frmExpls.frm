VERSION 5.00
Begin VB.Form frmExpls 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "For Next"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   3855
   End
End
Attribute VB_Name = "frmExpls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Find the 70th term of the series   1, 3, 5, 7,...

    'term = 1
    'For i = 1 To 69
        'term = term + 2
    'Next i


Private Sub Form_Load()
    term = 1
    
    For i = 1 To 70 Step 2
        term = i
        Text1.Text = term & ","
    Next i
End Sub

Private Sub Text1_Change()

End Sub
