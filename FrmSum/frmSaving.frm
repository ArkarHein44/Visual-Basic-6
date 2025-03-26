VERSION 5.00
Begin VB.Form frmSaving 
   Caption         =   "Sum Two Numbers"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSum 
      Caption         =   "Sum"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtSecondNo 
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtFirstNo 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label label3 
      Alignment       =   2  'Center
      Caption         =   "Result ="
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Second Number"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "First Number"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmSaving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSum_Click()
    Dim firstNo As Double
    Dim SecondNo As Double
    Dim Result As Double
    
    'Check if either textbox is empty
    If Trim(txtFirstNo.Text) = "" Or Trim(txtSecondNo.Text) = "" Then
        lblResult.Caption = "At least one number needed"
        Exit Sub
    End If
    
    'Validate that both inputs are numeric
    If Not IsNumeric(txtFirstNo.Text) Or Not IsNumeric(txtSecondNo.Text) Then
        lblResult.Caption = "Please enter valid numbers"
        Exit Sub
    End If
    
    'Convert the text values to numbers
    firstNo = CDbl(txtFirstNo.Text)
    SecondNo = Val(txtSecondNo.Text)
    
    'Calculate and display the result
    Result = firstNo + SecondNo
    lblResult.Caption = Format(Result)
End Sub

