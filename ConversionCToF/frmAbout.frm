VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About Conversation"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   2010
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "Conversion Application:" & vbCrLf & "This program changes Degrees" & vbCrLf & "in Centigrade to Fahrenheit" & vbCrLf & "and vice vasa."
End Sub
