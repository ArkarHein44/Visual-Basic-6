VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFormatTextCDC 
   Caption         =   "Format Selected Text"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1455
      Left            =   848
      TabIndex        =   2
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmFormatTextCDC.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Text"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
End
Attribute VB_Name = "frmFormatTextCDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComDlg()
    With CommonDialog1
        .Flags = .Flags Or cdlCFForceFontExist Or cdlCFEffects
    
        If IsNull(RichTextBox1.SelFontName) Then
            .Flags = .Flags Or cdlCFNoFaceSel
        Else
            .FontName = RichTextBox1.SelFontName
        End If
    
        If IsNull(RichTextBox1.SelFontSize) Then
            .Flags = .Flags Or cdlCFNoSizeSel
        Else
            .FontSize = RichTextBox1.SelFontSize
        End If
    
        If IsNull(RichTextBox1.SelBold) Or IsNull(RichTextBox1.SelItalic) Then
            .Flags = .Flags Or cdlCFNoStyleSel
        Else
            .FontBold = RichTextBox1.SelBold
            .FontItalic = RichTextBox1.SelItalic
            .Color = RichTextBox1.SelColor
        End If
    
        .CancelError = True
        .ShowFont
    
        If Err = 0 Then
            RichTextBox1.SelFontName = .FontName
            RichTextBox1.SelBold = .FontBold
            RichTextBox1.SelItalic = .FontItalic
            RichTextBox1.SelColor = .Color
            If (.Flags And cdlCFNoSizeSel) = 0 Then
                RichTextBox1.SelFontSize = .FontSize
            End If
            RichTextBox1.SelUnderline = .FontUnderline
            RichTextBox1.SelStrikeThru = .FontStrikethru
            RichTextBox1.SelColor = .Color
        End If
        End With
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdSelect_Click()
    On Error Resume Next
    Call ComDlg
End Sub

