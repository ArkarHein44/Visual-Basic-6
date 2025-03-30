VERSION 5.00
Begin VB.Form frmListBox 
   Caption         =   "Demo for List Boxes"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   2550
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdBackAll 
      Caption         =   "<<"
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
      Left            =   2670
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
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
      Left            =   2670
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdMoveAll 
      Caption         =   ">>"
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
      Left            =   2670
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">"
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
      Left            =   2670
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox lstRight 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   3990
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox lstLeft 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   270
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    Dim i As Integer
    If lstRight.SelCount > 0 Then
        i = lstRight.ListCount - 1
        
        Do While i >= 0
            If lstRight.Selected(i) = True Then
                lstLeft.AddItem lstRight.List(i)
                lstRight.RemoveItem (i)
            End If
            i = i - 1
        Loop
    End If
End Sub

Private Sub cmdBackAll_Click()
    Dim i As Integer
    For i = 0 To lstRight.ListCount - 1
        lstLeft.AddItem lstRight.List(i)
    Next i
    
    Do Until lstRight.ListCount = 0
        lstRight.RemoveItem (0)
    Loop
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click()
    Dim i As Integer
    If lstLeft.SelCount > 0 Then
        i = lstLeft.ListCount - 1
        
        Do While i >= 0
            If lstLeft.Selected(i) = True Then
                lstRight.AddItem lstLeft.List(i)
                lstLeft.RemoveItem i
            End If
            i = i - 1
        Loop
    End If
End Sub

Private Sub cmdMoveAll_Click()
    Dim i As Integer
    For i = 0 To lstLeft.ListCount - 1
        lstRight.AddItem lstLeft.List(i)
    Next i
    
    Do Until lstLeft.ListCount = 0
        lstLeft.RemoveItem (0)
    Loop
End Sub

Private Sub Form_Load()
    Dim i As Byte
    For i = 1 To 10
        lstLeft.AddItem "Line - " & CStr(i)
    Next i
End Sub
