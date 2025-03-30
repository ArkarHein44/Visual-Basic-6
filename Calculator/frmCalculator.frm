VERSION 5.00
Begin VB.Form frmCalculator 
   Caption         =   "Calculator"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalc 
      Caption         =   "="
      Height          =   495
      Left            =   6165
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cboOperator 
      Height          =   405
      Left            =   2205
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtSNum 
      Alignment       =   2  'Center
      Height          =   525
      Left            =   4245
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtFNum 
      Alignment       =   2  'Center
      Height          =   525
      Left            =   285
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7365
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCalc_Click()
    Dim tmp As Double
    If cboOperator.ListIndex = -1 Then
        MsgBox "Operator?", vbQuestion + vbOKOnly
        cboOperator.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtFNum.Text) = False Then
        MsgBox "Invalid Number.", vbExclamation + vbOKOnly
        txtFNum.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtSNum.Text) = False Then
        MsgBox "Invalid Number.", vbExclamation + vbOKOnly
        txtSNum.SetFocus
        Exit Sub
    End If
    
    tmp = CLng(txtSNum.Text)
    If cboOperator.Text = "/" And tmp = 0 Then
        MsgBox "Invalid Divisor", vbCritical + vbOKOnly
        txtSNum.SetFocus
        Exit Sub
    End If
    
    Select Case cboOperator.Text
        Case "+"
            lblResult = CStr(CLng(txtFNum) + CLng(txtSNum))
        Case "-"
            lblResult = CStr(CLng(txtFNum) - CLng(txtSNum))
        Case "*"
            lblResult = CStr(CLng(txtFNum) * CLng(txtSNum))
        Case "/"
            lblResult = CStr(CLng(txtFNum) / CLng(txtSNum))
    End Select
End Sub

Private Sub Form_Load()
    cboOperator.AddItem "+"
    cboOperator.AddItem "-"
    cboOperator.AddItem "*"
    cboOperator.AddItem "/"
End Sub
