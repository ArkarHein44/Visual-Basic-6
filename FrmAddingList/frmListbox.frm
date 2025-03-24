VERSION 5.00
Begin VB.Form frmListbox 
   Caption         =   "Adding Text to List Box"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddList 
      BackColor       =   &H8000000D&
      Caption         =   "Add List"
      Height          =   615
      Left            =   2100
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddList_Click()
    List1.AddItem Text1.Text
    Text1.SetFocus
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub
