VERSION 5.00
Begin VB.Form frmCalculateArea 
   Caption         =   "Calculation Area Using Public variable"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
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
      Left            =   1500
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtN2 
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
      Left            =   1980
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtN1 
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
      Left            =   1980
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Radius:"
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
      Left            =   1020
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Area:"
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
      Left            =   1020
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Formula = Pi * r^2"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmCalculateArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public area, radius As Double
Private pi As Double

Private Function CalArea(z As Double)
    pi = 4 * Atn(1)
    CalArea = pi * z * z
End Function

Public Sub Setup(y As Double)
    radius = y
    area = CalArea(radius)
End Sub

Private Sub cmdCalculate_Click()
    Dim X As Double
    X = Val(txtN1.Text)
    Call Setup(X)
    txtN2.Text = Format(area, ####,####")
End Sub
