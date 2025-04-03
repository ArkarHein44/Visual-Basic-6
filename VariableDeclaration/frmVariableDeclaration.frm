VERSION 5.00
Begin VB.Form frmVariableDeclaration 
   Caption         =   "Variable Declaration"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   2985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
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
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
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
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
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
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      Left            =   600
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmVariableDeclaration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Temp As Integer
Public num As Integer

Sub Test()
    Dim Temp As Integer
    Temp = 200
    MsgBox "Temp is " & Form1.Temp
End Sub

Private Sub Command1_Click()
    Test
End Sub

Private Sub Form_Load()
    Temp = 50
End Sub

Private Sub cmdTotal_Click()
    num = Val(Text1.Text)
    Text2.Text = Str(RunningTotal(num))
End Sub

Function RunningTotal(num)
    Static Temp 'static stored value
    Temp = Temp + num
    RunningTotal = Temp
End Function


