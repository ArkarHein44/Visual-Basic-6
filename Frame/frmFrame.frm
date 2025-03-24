VERSION 5.00
Begin VB.Form frmFrame 
   Caption         =   "Using Frame"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   255
      Width           =   5175
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   1245
         TabIndex        =   4
         Top             =   3360
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1935
         Left            =   2520
         Picture         =   "frmFrame.frx":0000
         ScaleHeight     =   1875
         ScaleWidth      =   2355
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label2 
         Height          =   735
         Left            =   720
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Height          =   615
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Picture1.Picture = LoadPicture("C:\Users\acer\Desktop\VB6\Frame/convo.jpg")
    Label1.Caption = Date
    Label2.Caption = Time
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii > 47 And KeyAscii <= 58 Then
        KeyAscii = 8
    End If
End Sub
