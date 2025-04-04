VERSION 5.00
Begin VB.Form frmConversion 
   Caption         =   "Temperature Conversion"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   3390
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
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
      Height          =   495
      Left            =   2310
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdFToC 
      Caption         =   "F To C"
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
      Left            =   1230
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCToF 
      Caption         =   "C To F"
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
      Left            =   150
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtF 
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
      Left            =   2130
      TabIndex        =   2
      Top             =   960
      Width           =   1815
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
      Left            =   2130
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Fahrenheit ="
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
      Left            =   690
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Centigrate ="
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
      Left            =   690
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variable declaration
Option Explicit
Dim C As Double
Dim F As Double

Private Sub cmdCToF_Click()
    C = CDbl(Val(txtC.Text))
    'Checking numeric input value of C
    If IsNumeric(txtC.Text) Then
        F = (C * 9 / 5) + 32 'C To F Formula
    txtF.Text = F
    Else
        MsgBox "Only number allow!!!", vbCritical + vbOKOnly, "Invalid input"
        txtC.Text = "" 'Clear C
        txtC.SetFocus 'on focus C
    End If
End Sub

Private Sub cmdFToC_Click()
    F = CDbl(Val(txtF.Text))
    'Checking numeric input value of F
    If IsNumeric(txtF.Text) Then
        C = (F - 32) * 5 / 9 'F To C Formula
        txtC.Text = C
    Else
        MsgBox "Only number allow!!!", vbCritical + vbOKOnly, "Invalid input"
        txtF.Text = "" 'Clear F
        txtF.SetFocus 'on focus F
    End If
End Sub

Private Sub cmdClear_Click()
    txtC.Text = ""
    txtF.Text = ""
End Sub

Private Sub cmdClose_Click()
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Are you sure you want to Exit?", vbQuestion + vbOKCancel, "Confirm Exist")
    If confirm = vbOK Then
        End
    End If
End Sub

