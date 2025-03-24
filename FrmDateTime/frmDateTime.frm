VERSION 5.00
Begin VB.Form frmDateTime 
   Caption         =   "Date & Time Program"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCLose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   1553
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Height          =   495
      Left            =   2393
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblDate 
      Height          =   495
      Left            =   2393
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label label2 
      Caption         =   "Time"
      Height          =   495
      Left            =   953
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label label1 
      Caption         =   "Date"
      Height          =   495
      Left            =   953
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCLose_Click()
    End
End Sub

Private Sub Form_Load()
    lblDate.Caption = date
    lblTime.Caption = time
End Sub

