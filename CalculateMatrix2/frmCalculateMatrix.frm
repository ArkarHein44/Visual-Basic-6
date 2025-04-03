VERSION 5.00
Begin VB.Form frmCalculateMatrix 
   Caption         =   "Calculate Matrix"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBeta 
      Alignment       =   2  'Center
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
      Left            =   4717
      TabIndex        =   37
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtAlpha 
      Alignment       =   2  'Center
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
      Left            =   1237
      TabIndex        =   34
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   4110
      TabIndex        =   31
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
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
      Left            =   2190
      TabIndex        =   30
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txt11c 
      Alignment       =   2  'Center
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
      Left            =   3570
      TabIndex        =   28
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt12c 
      Alignment       =   2  'Center
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
      Left            =   4050
      TabIndex        =   27
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt13c 
      Alignment       =   2  'Center
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
      Left            =   4530
      TabIndex        =   26
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt21c 
      Alignment       =   2  'Center
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
      Left            =   3570
      TabIndex        =   25
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txt22c 
      Alignment       =   2  'Center
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
      Left            =   4050
      TabIndex        =   24
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txt23c 
      Alignment       =   2  'Center
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
      Left            =   4530
      TabIndex        =   23
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txt31c 
      Alignment       =   2  'Center
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
      Left            =   3570
      TabIndex        =   22
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txt32c 
      Alignment       =   2  'Center
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
      Left            =   4050
      TabIndex        =   21
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txt33c 
      Alignment       =   2  'Center
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
      Left            =   4530
      TabIndex        =   20
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txt33b 
      Alignment       =   2  'Center
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
      Left            =   6742
      TabIndex        =   18
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txt32b 
      Alignment       =   2  'Center
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
      Left            =   6262
      TabIndex        =   17
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txt31b 
      Alignment       =   2  'Center
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
      Left            =   5782
      TabIndex        =   16
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txt23b 
      Alignment       =   2  'Center
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
      Left            =   6742
      TabIndex        =   15
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txt22b 
      Alignment       =   2  'Center
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
      Left            =   6262
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txt21b 
      Alignment       =   2  'Center
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
      Left            =   5782
      TabIndex        =   13
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txt13b 
      Alignment       =   2  'Center
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
      Left            =   6742
      TabIndex        =   12
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txt12b 
      Alignment       =   2  'Center
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
      Left            =   6262
      TabIndex        =   11
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txt11b 
      Alignment       =   2  'Center
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
      Left            =   5782
      TabIndex        =   10
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txt33a 
      Alignment       =   2  'Center
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
      Left            =   3262
      TabIndex        =   8
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txt32a 
      Alignment       =   2  'Center
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
      Left            =   2782
      TabIndex        =   7
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txt31a 
      Alignment       =   2  'Center
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
      Left            =   2302
      TabIndex        =   6
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txt23a 
      Alignment       =   2  'Center
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
      Left            =   3262
      TabIndex        =   5
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txt22a 
      Alignment       =   2  'Center
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
      Left            =   2782
      TabIndex        =   4
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txt21a 
      Alignment       =   2  'Center
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
      Left            =   2302
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txt13a 
      Alignment       =   2  'Center
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
      Left            =   3262
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txt12a 
      Alignment       =   2  'Center
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
      Left            =   2782
      TabIndex        =   1
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txt11a 
      Alignment       =   2  'Center
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
      Left            =   2302
      TabIndex        =   0
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Beta"
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
      Left            =   4597
      TabIndex        =   39
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "*"
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
      Left            =   5197
      TabIndex        =   38
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "*"
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
      Left            =   1717
      TabIndex        =   36
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Alpha"
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
      Left            =   1117
      TabIndex        =   35
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "C=Alpha*A + Beta*B"
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
      Left            =   2730
      TabIndex        =   33
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "A = [aij], B = [bij]"
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
      Left            =   3030
      TabIndex        =   32
      Top             =   360
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   3742
      X2              =   3742
      Y1              =   1560
      Y2              =   3120
   End
   Begin VB.Line Line5 
      X1              =   3622
      X2              =   3742
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line7 
      X1              =   5662
      X2              =   5662
      Y1              =   1560
      Y2              =   3120
   End
   Begin VB.Line Line8 
      X1              =   5662
      X2              =   5782
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line18 
      X1              =   3450
      X2              =   3450
      Y1              =   3600
      Y2              =   5160
   End
   Begin VB.Line Line17 
      X1              =   3450
      X2              =   3570
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line16 
      X1              =   3450
      X2              =   3570
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line15 
      X1              =   5010
      X2              =   5010
      Y1              =   3600
      Y2              =   5160
   End
   Begin VB.Line Line14 
      X1              =   4890
      X2              =   5010
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line13 
      X1              =   4890
      X2              =   5010
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "C= Alpha*(A)+Beta*(B)="
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
      Left            =   600
      TabIndex        =   29
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "B="
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
      Left            =   4237
      TabIndex        =   19
      Top             =   2160
      Width           =   495
   End
   Begin VB.Line Line12 
      X1              =   7102
      X2              =   7222
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line11 
      X1              =   7102
      X2              =   7222
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line10 
      X1              =   7222
      X2              =   7222
      Y1              =   1560
      Y2              =   3120
   End
   Begin VB.Line Line9 
      X1              =   5662
      X2              =   5782
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "A="
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
      Left            =   637
      TabIndex        =   9
      Top             =   2280
      Width           =   495
   End
   Begin VB.Line Line6 
      X1              =   3622
      X2              =   3742
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   2182
      X2              =   2302
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      X1              =   2182
      X2              =   2302
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   2182
      X2              =   2182
      Y1              =   1560
      Y2              =   3120
   End
End
Attribute VB_Name = "frmCalculateMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j, k As Integer
Dim a(1 To 3, 1 To 3) As Double
Dim b(1 To 3, 1 To 3) As Double
Dim c(1 To 3, 1 To 3) As Double
Dim Alpha, Beta As Variant

Private Sub cmdCalculate_Click()
    For i = 1 To 3
        For j = 1 To 3
            a(i, j) = Val(InputBox(k))
        Next j
    Next i
    
    txt11a.Text = a(1, 1)
    txt12a.Text = a(1, 2)
    txt13a.Text = a(1, 3)
    txt21a.Text = a(2, 1)
    txt22a.Text = a(2, 2)
    txt23a.Text = a(2, 3)
    txt31a.Text = a(3, 1)
    txt32a.Text = a(3, 2)
    txt33a.Text = a(3, 3)
    
    For j = 1 To 3
        For i = 1 To 3
            b(i, j) = Val(InputBox(k))
        Next i
    Next j
    txt11b.Text = b(1, 1)
    txt12b.Text = b(1, 2)
    txt13b.Text = b(1, 3)
    txt21b.Text = b(2, 1)
    txt22b.Text = b(2, 2)
    txt23b.Text = b(2, 3)
    txt31b.Text = b(3, 1)
    txt32b.Text = b(3, 2)
    txt33b.Text = b(3, 3)
    
    Alpha = Val(txtAlpha.Text)
    Beta = Val(txtBeta.Text)
    
    For j = 1 To 3
        For i = 1 To 3
            c(i, j) = Alpha * a(i, j) + Beta * b(i, j)
            MsgBox (c(i, j))
        Next i
    Next j
    txt11c.Text = c(1, 1)
    txt12c.Text = c(1, 2)
    txt13c.Text = c(1, 3)
    txt21c.Text = c(2, 1)
    txt22c.Text = c(2, 2)
    txt23c.Text = c(2, 3)
    txt31c.Text = c(3, 1)
    txt32c.Text = c(3, 2)
    txt33c.Text = c(3, 3)
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

