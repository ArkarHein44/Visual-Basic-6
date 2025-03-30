VERSION 5.00
Begin VB.Form frmCheckBox 
   Caption         =   "CheckBox Program"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   5025
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
      Height          =   495
      Left            =   3465
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
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
      Left            =   2145
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
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
      Height          =   1575
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.CheckBox chkLower 
      Caption         =   "Lower  Case"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CheckBox chkUpper 
      Caption         =   "Upper Case"
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
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Case"
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
      Left            =   345
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a setence"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
    If chkUpper.Value = 1 And chkLower.Value = 0 Then
        Text1 = Format(Text1.Text, ">")
    ElseIf chkLower.Value = 1 And chkUpper.Value = 0 Then
        Text1 = Format(Text1.Text, "<")
    ElseIf chkUpper.Value = 1 And chkLower.Value = 1 Then
        MsgBox "Please, Choose only one case to change!", vbInformation + vbOKOnly, "Message BOx"
    Else
        MsgBox "Please, Choose one case to change!", vbCritical + vbOKCancel, "Message Box"
    End If
End Sub

Private Sub cmdClear_Click()
    Text1.Text = ""
End Sub

Private Sub cmdExit_Click()
    End
End Sub
