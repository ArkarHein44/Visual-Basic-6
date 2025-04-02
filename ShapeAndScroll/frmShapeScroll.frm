VERSION 5.00
Begin VB.Form frmShapeScroll 
   Caption         =   "Shape and Scroll Bars Demo"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
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
      Left            =   2430
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin VB.OptionButton optRBox 
      Caption         =   "Rouinded Box"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.OptionButton optCircle 
      Caption         =   "Circle"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton optBox 
      Caption         =   "Box"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox chkFill 
      Caption         =   "Fill"
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
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   3615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      Left            =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   120
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmShapeScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFill_Click()
    If chkFill.Value = vbChecked Then
        Shape1.FillStyle = 0
    Else
        Shape1.FillStyle = 1
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub HScroll1_Change()
    Shape1.Width = HScroll1.Value
End Sub

Private Sub optBox_Click()
    If optBox.Value Then
        Shape1.Shape = 0
    End If
End Sub

Private Sub optCircle_Click()
    If optCircle.Value Then
        Shape1.Shape = 2
    End If
End Sub

Private Sub optRBox_Click()
    If optRBox.Value Then
        Shape1.Shape = 4
    End If
    
End Sub

Private Sub VScroll1_Change()
    Shape1.Height = VScroll1.Value
End Sub
