VERSION 5.00
Begin VB.Form frmOptProgram 
   Caption         =   "Option Program"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6510
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
   ScaleHeight     =   5265
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   1208
      TabIndex        =   9
      Top             =   4440
      Width           =   4095
   End
   Begin VB.CommandButton cmdShowMsgBox 
      Caption         =   "Show Message Box"
      Height          =   495
      Left            =   1208
      TabIndex        =   8
      Top             =   3840
      Width           =   4095
   End
   Begin VB.OptionButton optYesNo 
      Caption         =   "Yes or No"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.OptionButton optOKCancel 
      Caption         =   "OK Cancel"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.OptionButton optOKOnly 
      Caption         =   "OK Only"
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
      Left            =   3960
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Button Group"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3555
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Frame fraSelectCase 
      Caption         =   "Select Case"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   675
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      Begin VB.OptionButton optExclam 
         Caption         =   "Exclamation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton optInfo 
         Caption         =   "Information"
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
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   608
      TabIndex        =   7
      Top             =   3240
      Width           =   5295
   End
End
Attribute VB_Name = "frmOptProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdShowMsgBox_Click()
    Dim K, IconType, ButtonGroup As Integer
    
    If optInfo.Value = True Then
        IconType = vbInformation
    ElseIf optExclam.Value = True Then
        IconType = vbExclamation
    End If
    
    If optOKOnly.Value = True Then
        ButtonGroup = vbOKOnly
    ElseIf optOKCancel.Value = True Then
        ButtonGroup = vbOKCancel
    ElseIf optYesNo.Value = True Then
        ButtonGroup = vbYesNo
    End If
    
    K = MsgBox("Welcome to VB6.0!", IconType + ButtonGroup, "Option Testing")
    
    Select Case K
        Case vbOK
            lblInfo.Caption = "You have clicked OK button"
        Case vbCancel
            lblInfo.Caption = "You have clicked Cancel button"
        Case vbYes
            lblInfo.Caption = "You have clicked Yes button"
        Case vbNo
            lblInfo.Caption = "You have clicked No button"
    End Select
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub
