VERSION 5.00
Begin VB.Form frmCalculateMatrix 
   Caption         =   "Calculate Matrix"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
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
      Left            =   3585
      TabIndex        =   31
      Top             =   5160
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
      Left            =   1545
      TabIndex        =   30
      Top             =   5160
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
      Left            =   2985
      TabIndex        =   28
      Top             =   3240
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
      Left            =   3465
      TabIndex        =   27
      Top             =   3240
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
      Left            =   3945
      TabIndex        =   26
      Top             =   3240
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
      Left            =   2985
      TabIndex        =   25
      Top             =   3720
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
      Left            =   3465
      TabIndex        =   24
      Top             =   3720
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
      Left            =   3945
      TabIndex        =   23
      Top             =   3720
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
      Left            =   2985
      TabIndex        =   22
      Top             =   4200
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
      Left            =   3465
      TabIndex        =   21
      Top             =   4200
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
      Left            =   3945
      TabIndex        =   20
      Top             =   4200
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
      Left            =   5505
      TabIndex        =   18
      Top             =   2400
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
      Left            =   5025
      TabIndex        =   17
      Top             =   2400
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
      Left            =   4545
      TabIndex        =   16
      Top             =   2400
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
      Left            =   5505
      TabIndex        =   15
      Top             =   1920
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
      Left            =   5025
      TabIndex        =   14
      Top             =   1920
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
      Left            =   4545
      TabIndex        =   13
      Top             =   1920
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
      Left            =   5505
      TabIndex        =   12
      Top             =   1440
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
      Left            =   5025
      TabIndex        =   11
      Top             =   1440
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
      Left            =   4545
      TabIndex        =   10
      Top             =   1440
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
      Left            =   2385
      TabIndex        =   8
      Top             =   2400
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
      Left            =   1905
      TabIndex        =   7
      Top             =   2400
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
      Left            =   1425
      TabIndex        =   6
      Top             =   2400
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
      Left            =   2385
      TabIndex        =   5
      Top             =   1920
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
      Left            =   1905
      TabIndex        =   4
      Top             =   1920
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
      Left            =   1425
      TabIndex        =   3
      Top             =   1920
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
      Left            =   2385
      TabIndex        =   2
      Top             =   1440
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
      Left            =   1905
      TabIndex        =   1
      Top             =   1440
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
      Left            =   1425
      TabIndex        =   0
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "C = A+B"
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
      Left            =   2760
      TabIndex        =   33
      Top             =   600
      Width           =   1095
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
      Left            =   2445
      TabIndex        =   32
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   2865
      X2              =   2865
      Y1              =   1320
      Y2              =   2880
   End
   Begin VB.Line Line5 
      X1              =   2745
      X2              =   2865
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line7 
      X1              =   4425
      X2              =   4425
      Y1              =   1320
      Y2              =   2880
   End
   Begin VB.Line Line8 
      X1              =   4425
      X2              =   4545
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line18 
      X1              =   2865
      X2              =   2865
      Y1              =   3120
      Y2              =   4680
   End
   Begin VB.Line Line17 
      X1              =   2865
      X2              =   2985
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line16 
      X1              =   2865
      X2              =   2985
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line15 
      X1              =   4425
      X2              =   4425
      Y1              =   3120
      Y2              =   4680
   End
   Begin VB.Line Line14 
      X1              =   4305
      X2              =   4425
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line13 
      X1              =   4305
      X2              =   4425
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "C="
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
      Left            =   2265
      TabIndex        =   29
      Top             =   3720
      Width           =   495
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
      Left            =   3825
      TabIndex        =   19
      Top             =   1920
      Width           =   495
   End
   Begin VB.Line Line12 
      X1              =   5865
      X2              =   5985
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line11 
      X1              =   5865
      X2              =   5985
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line10 
      X1              =   5985
      X2              =   5985
      Y1              =   1320
      Y2              =   2880
   End
   Begin VB.Line Line9 
      X1              =   4425
      X2              =   4545
      Y1              =   2880
      Y2              =   2880
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
      Left            =   705
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.Line Line6 
      X1              =   2745
      X2              =   2865
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      X1              =   1305
      X2              =   1425
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   1305
      X2              =   1425
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   1305
      X2              =   1305
      Y1              =   1320
      Y2              =   2880
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

Private Sub cmdCalculate_Click()
    For i = 1 To 3
        For j = 1 To 3
            a(i, j) = Val(InputBox(k))
            txt11a.Text = a(1, 1)
            txt12a.Text = a(1, 2)
            txt13a.Text = a(1, 3)
            txt21a.Text = a(2, 1)
            txt22a.Text = a(2, 2)
            txt23a.Text = a(2, 3)
            txt31a.Text = a(3, 1)
            txt32a.Text = a(3, 2)
            txt33a.Text = a(3, 3)
        Next j
    Next i
    
    For i = 1 To 3
        For j = 1 To 3
            b(i, j) = Val(InputBox(k))
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
    Next i
    
    For j = 1 To 3
        For i = 1 To 3
            c(i, j) = a(i, j) + b(i, j)
            MsgBox (c(i, j))
        Next i
        txt11c.Text = c(1, 1)
        txt12c.Text = c(1, 2)
        txt13c.Text = c(1, 3)
        txt21c.Text = c(2, 1)
        txt22c.Text = c(2, 2)
        txt23c.Text = c(2, 3)
        txt31c.Text = c(3, 1)
        txt32c.Text = c(3, 2)
        txt33c.Text = c(3, 3)
    Next j
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub
