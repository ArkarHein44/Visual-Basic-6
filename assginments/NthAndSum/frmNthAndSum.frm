VERSION 5.00
Begin VB.Form frmNthAndSum 
   Caption         =   "Finding n-th Term and Total Sum"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
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
      Left            =   3780
      TabIndex        =   11
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSum 
      Caption         =   "Find &Total Sum"
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
      Left            =   1980
      TabIndex        =   10
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdNth 
      Caption         =   "Fint &n-th Term"
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
      Left            =   180
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtD 
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
      Left            =   900
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtN 
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
      Left            =   900
      TabIndex        =   3
      Top             =   1080
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
      Left            =   900
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblSum 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3720
      TabIndex        =   13
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblNth 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3720
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Total Sum ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "n_th Term ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Enter value of a, n and d"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2580
      TabIndex        =   6
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "d = "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "n = "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "a = "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "frmNthAndSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Long
Dim n As Long
Dim d As Long
Dim nth As Long
Dim Sum As Double

Private Sub ClearInput()
    txtA.Text = ""
    txtN.Text = ""
    txtD.Text = ""
    txtA.SetFocus
End Sub

Private Sub cmdClear_Click()
    ClearInput
    lblNth.Caption = ""
    lblSum.Caption = ""
End Sub

Private Sub cmdNth_Click()
    If IsNumeric(txtA.Text) And IsNumeric(txtN.Text) And IsNumeric(txtD.Text) Then
        a = CLng(Val(txtA.Text))
        n = CLng(Val(txtN.Text))
        d = CLng(Val(txtD.Text))
        
        nth = a + (n - 1) * d
        lblNth.Caption = nth
        
    Else
        MsgBox "Only number Allow!!!", vbExclamation + vbOKOnly, "Invalid input"
        ClearInput
    End If
End Sub

Private Sub cmdSum_Click()
    If IsNumeric(txtA.Text) And IsNumeric(txtN.Text) And IsNumeric(txtD.Text) Then
        a = CLng(Val(txtA.Text))
        n = CLng(Val(txtN.Text))
        d = CLng(Val(txtD.Text))
        
        Sum = n / 2 * (2 * a + (n - 1) * d)
        lblSum.Caption = Sum
        
    Else
        MsgBox "Only number Allow!!!", vbExclamation + vbOKOnly, "Invalid input"
        ClearInput
    End If
End Sub
