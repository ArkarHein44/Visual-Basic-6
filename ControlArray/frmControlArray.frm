VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFields 
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
      Left            =   2880
      TabIndex        =   7
      Text            =   "qrstuvwxyz"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
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
      Left            =   2880
      TabIndex        =   6
      Text            =   "hijklmnop"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
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
      Left            =   2880
      TabIndex        =   5
      Text            =   "abcdefg"
      Top             =   240
      Width           =   1815
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
      Height          =   405
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton optFrequency 
      Caption         =   "Option4"
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
      Index           =   3
      Left            =   5400
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.OptionButton optFrequency 
      Caption         =   "Option3"
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
      Index           =   2
      Left            =   5400
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton optFrequency 
      Caption         =   "Option2"
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
      Index           =   1
      Left            =   5400
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.OptionButton optFrequency 
      Caption         =   "Option1"
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
      Index           =   0
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module-level variable
Dim optFrequencyIndex As Integer

Private Sub optFrequency_Click(Index As Integer)
    'Remeber the last button selected
    optFrequencyIndex = Index
End Sub

Private Sub LoadText1()
    ' Suppose you created Text(0) at design time.
    Load Text1(1)
    Load Text1(2)
    Load Text1(3)
    
    Text1(0).Move 1200, 1500, 500, 350
    Text1(1).Move 1200, 2000, 800, 350
    Text1(2).Move 1200, 2500, 1500, 350
    Text1(3).Move 1200, 3000, 1000, 350 'Remark: text1(0).Move (left,top, width,height)
    
    'Text1(1).Move 1200, 2000, 800, 350      ' Set other properties as required.
    Text1(1).MaxLength = 10         ' Finally make it visible.
    Text1(1).Visible = True
    
    Text1(2).MaxLength = 10         ' Finally make it visible.
    Text1(2).Visible = True
    
    Text1(3).MaxLength = 10         ' Finally make it visible.
    Text1(3).Visible = True
End Sub

Private Sub Form_Load()
    LoadText1
    
    
End Sub

Private Sub txtFields_Change(Index As Integer)
    For i = txtFields.LBound To txtFields.UBound
        txtFields(i).Text = ""
    Next
End Sub
