VERSION 5.00
Begin VB.Form frmView 
   Caption         =   "Image Viewer"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3668
      TabIndex        =   4
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   615
      Left            =   1388
      TabIndex        =   3
      Top             =   5520
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   405
      Left            =   315
      TabIndex        =   2
      Top             =   3720
      Width           =   1920
   End
   Begin VB.DirListBox Dir1 
      Height          =   1350
      Left            =   2460
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   1230
      Left            =   4680
      Pattern         =   """*.jpg,*.bmp,*.icon"""
      TabIndex        =   0
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   1748
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Directory"
      Height          =   255
      Left            =   2580
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "File List Box"
      Height          =   255
      Left            =   4740
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Drive"
      Height          =   255
      Left            =   420
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdShow_Click()
    Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Me.File1.Pattern = "*.jpg;*.bmp;*.icon;*.png"
End Sub
