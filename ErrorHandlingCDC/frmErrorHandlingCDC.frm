VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmErrorHandlingCDC 
   Caption         =   "Error Handling using common dialog control"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   2160
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
      Top             =   2160
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
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmErrorHandlingCDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDisplay_Click()
    On Error Resume Next
        CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
        CommonDialog1.ShowFont
        Text1.FontName = CommonDialog1.FontName
        Text1.FontSize = CommonDialog1.FontName 'Error Occur, it will go to Next
        Text1.FontBold = CommonDialog1.FontBold
        Text1.FontItalic = CommonDialog1.FontItalic
        Text1.FontUnderline = CommonDialog1.FontSize
        Text1.FontStrikethru = CommonDialog1.FontStrikethru
        Text1.ForeColor = CommonDialog1.Color
End Sub
