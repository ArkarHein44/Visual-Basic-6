VERSION 5.00
Begin VB.Form frmChanging 
   Caption         =   "Changing  Caption to Command Button"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChange 
      Caption         =   "Changing"
      Height          =   495
      Left            =   833
      TabIndex        =   1
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   833
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "frmChanging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
    cmdChange.Caption = Text1.Text
End Sub

