VERSION 5.00
Begin VB.Form frmDiscriminant 
   BackColor       =   &H00FF8080&
   Caption         =   "Discriminant Calculator"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4065
      TabIndex        =   12
      Top             =   3360
      Width           =   1800
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFFF&
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
      Height          =   615
      Left            =   2145
      TabIndex        =   11
      Top             =   3360
      Width           =   1800
   End
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
      Left            =   225
      TabIndex        =   9
      Top             =   3360
      Width           =   1800
   End
   Begin VB.TextBox txtResult2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   5475
   End
   Begin VB.TextBox txtResult1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   5475
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   5475
   End
   Begin VB.TextBox txtC 
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
      Left            =   4538
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtB 
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
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtA 
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
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Discriminant calculator  (b ^2- 4ac)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "frmDiscriminant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- variable declaration section---
Option Explicit
Dim a As Double
Dim b As Double
Dim c As Double
Dim del As Double
Dim X1 As Double
Dim X2 As Double
'----------------------------------

Private Sub cmdCalculate_Click()

    '--checking click cmdCalculate without data---
    If txtA.Text = "" Or txtB.Text = "" Or txtC.Text = "" Then
        MsgBox "a,b and c can't be empty"
        
    '--Checking numeric value from text boxes ----
    ElseIf Not IsNumeric(txtA.Text) Or Not IsNumeric(txtB.Text) Or Not IsNumeric(txtC.Text) Then
        MsgBox "Only numbers allow"
    Else
    'Calculation and logic
        a = CDbl(txtA.Text)
        b = CDbl(txtB.Text)
        c = CDbl(txtC.Text)
        
        'main formula
        del = b * b - 4 * a * c
        
        If del = 0 Then
            txtDisplay.Text = "Equal roots"
            X1 = 0.5 * (-b) / a
            txtResult1.Text = "X1 = " & Format(X1, "Fixed")
            txtResult2.Text = "X2 = " & Format(X1, "Fixed")
        ElseIf del > 0 Then
            txtDisplay.Text = "Real Unequal roots"
            X1 = 0.5 * (-b + Sqr(del)) / a
            X2 = 0.5 * (-b - Sqr(del)) / a
            txtResult1.Text = "X1 = " & Format(X1, "Fixed")
            txtResult2.Text = "X2 = " & Format(X2, "Fixed")
        Else
            txtDisplay.Text = "Complex roots"
        End If
    End If
End Sub

Private Sub cmdClear_Click()
    txtA.Text = ""
    txtB.Text = ""
    txtC.Text = ""
    txtDisplay.Text = ""
    txtResult1.Text = ""
    txtResult2.Text = ""
End Sub

Private Sub cmdExit_Click()
    End
End Sub

