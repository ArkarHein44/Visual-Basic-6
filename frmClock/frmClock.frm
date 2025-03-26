VERSION 5.00
Begin VB.Form frmClock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clock Program"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2745
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   945
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   2400
      Top             =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   225
      X2              =   5145
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   285
      TabIndex        =   0
      Top             =   720
      Width           =   4815
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrTimer_Timer()
    lblTime.Caption = Time
End Sub

Private Sub cmdPause_Click()
    tmrTimer.Enabled = True
    cmdPause.Enabled = False
    cmdResume.Enabled = True
End Sub

Private Sub cmdResume_Click()
    tmrTimer.Enabled = True
    cmdPause.Enabled = True
    cmdResume.Enabled = False
End Sub

Private Sub Load()
    lblTime.Caption = Time
    cmdResume.Enabled = False
End Sub
