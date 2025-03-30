VERSION 5.00
Begin VB.Form frmListBox2 
   Caption         =   "List Box 2"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6285
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
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ListBox lstCollection 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   3255
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtToInput 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   375
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmListBox2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdAdd_Click()
    lstCollection.AddItem (txtToInput.Text)
    txtToInput.Text = ""
    txtToInput.SetFocus
End Sub

Private Sub cmdClear_Click()
    lstCollection.Clear
End Sub

Private Sub cmdDelete_Click()
    lstCollection.RemoveItem (i)
End Sub

