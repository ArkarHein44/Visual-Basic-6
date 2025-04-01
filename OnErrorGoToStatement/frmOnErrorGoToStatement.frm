VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOnErrorGoTo 
   Caption         =   "On Error Go To"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "&Display"
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
      Left            =   1553
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   473
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Type any text  in the Text Box"
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
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmOnErrorGoTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDisplay_Click()
    On Error GoTo 0
        CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
        CommonDialog1.ShowFont
        Text1.FontName = CommonDialog1.FontSize 'Error Occur, it will pass the eror and go to statement
        Text1.FontSize = CommonDialog1.FontSize
        Text1.FontBold = CommonDialog1.FontBold
        Text1.FontItalic = CommonDialog1.FontItalic
        Text1.FontUnderline = CommonDialog1.FontSize
        Text1.FontStrikethru = CommonDialog1.FontStrikethru
        Text1.ForeColor = CommonDialog1.Color
ErrHandler:
    MsgBox "Error!", vbCritical + vbOKOnly, "Error Handling"
End Sub

