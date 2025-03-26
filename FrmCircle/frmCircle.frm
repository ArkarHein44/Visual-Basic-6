VERSION 5.00
Begin VB.Form frmCircle 
   AutoRedraw      =   -1  'True
   Caption         =   "To draw circles by Hsb"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbRadius 
      Height          =   375
      Left            =   353
      Max             =   1250
      Min             =   1
      TabIndex        =   0
      Top             =   1800
      Value           =   500
      Width           =   3855
   End
   Begin VB.Label lblArea 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Area Now ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "frmCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- Declaration Section ----
Option Explicit
Const PI As Single = 3.14159
Dim Radius As Single
'----------------------------

Private Sub hsbRadius_Change()
    Dim CY As Long
    Dim CX As Long
    
    CX = Me.ScaleWidth \ 2
    CY = (Me.ScaleHeight - 600) \ 2
    
    Radius = hsbRadius.Value
    Me.Cls
    Me.Circle (CX, CY), Radius 'To Draw circle, center and radius'
    lblArea.Caption = CStr(Radius * Radius * PI)
        
End Sub
Private Sub hsbRadius_Scroll()
    Call hsbRadius_Change
End Sub

Private Sub Form_Activate()
    Call hsbRadius_Change
End Sub
