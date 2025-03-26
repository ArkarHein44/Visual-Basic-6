VERSION 5.00
Begin VB.Form value_of_Z 
   Caption         =   "value_of_z"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
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
      Height          =   375
      Left            =   3645
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
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
      Height          =   375
      Left            =   2205
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
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
      Height          =   375
      Left            =   765
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   3247
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
      Begin VB.TextBox txtZ 
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
         Left            =   720
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "z="
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Each Value"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   322
      TabIndex        =   0
      Top             =   840
      Width           =   2895
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
         Height          =   400
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
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
         Height          =   400
         Left            =   1200
         TabIndex        =   11
         Top             =   960
         Width           =   1455
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
         Height          =   400
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "c="
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
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "b="
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
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "a="
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
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Caption         =   "The value of z = 2a+3b-4c."
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
      Left            =   105
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "value_of_Z"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --- Declaration Section ----
Option Explicit
Dim A As Double
Dim B As Double
Dim C As Double
Dim Z As Double
' ----------------------------

Private Sub cmdCalculate_Click()
        
    If txtA.Text = "" Or txtB.Text = "" Or txtC.Text = "" Then
        MsgBox "a, b and c are cannot be empty!!!"
    ElseIf Not IsNumeric(txtA.Text) Or Not IsNumeric(txtB.Text) Or Not IsNumeric(txtC.Text) Then
        MsgBox "Allow numbers only!!!"
        txtA.Text = ""
        txtB.Text = ""
        txtC.Text = ""
        txtZ.Text = ""
    Else
        A = CDbl(txtA.Text)
        B = CDbl(txtB.Text)
        C = CDbl(txtC.Text)
        Z = (2 * A) + (3 * B) - (4 * C)
        txtZ.Text = Format(Z)
    End If
    
End Sub

Private Sub cmdClear_Click()
    txtA.Text = ""
    txtB.Text = ""
    txtC.Text = ""
    txtZ.Text = ""
End Sub

Private Sub cmdExit_Click()
    End
End Sub

