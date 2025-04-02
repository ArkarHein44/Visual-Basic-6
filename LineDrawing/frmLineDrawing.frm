VERSION 5.00
Begin VB.Form frmLineDrawing 
   Caption         =   "Line Drawing"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
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
      Left            =   3488
      TabIndex        =   9
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
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
      Left            =   1568
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtY2 
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
      Left            =   4268
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtX2 
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
      Left            =   1395
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtY1 
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
      Left            =   4268
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtX1 
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
      Left            =   1395
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Y2 ="
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
      Left            =   3428
      TabIndex        =   7
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "X2 ="
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
      Left            =   555
      TabIndex        =   5
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Y1 ="
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
      Left            =   3428
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "X1 ="
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
      Left            =   555
      TabIndex        =   1
      Top             =   3360
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   1208
      X2              =   5183
      Y1              =   240
      Y2              =   2055
   End
End
Attribute VB_Name = "frmLineDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDraw_Click()
    Line1.X1 = Val(txtX1.Text)
    Line1.Y1 = Val(txtY1.Text)
    Line1.X2 = Val(txtX2.Text)
    Line1.Y2 = Val(txtY2.Text)
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

