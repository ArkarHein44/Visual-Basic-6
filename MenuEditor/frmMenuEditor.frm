VERSION 5.00
Begin VB.Form frmShapeColorMenu 
   Caption         =   "Menu Editor"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu Hyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu mnuFigure 
      Caption         =   "&Figure"
      Begin VB.Menu mnuRectangle 
         Caption         =   "&Rectangle"
      End
      Begin VB.Menu mnuCircle 
         Caption         =   "&Circle"
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "&Selection"
      Begin VB.Menu mnuColors 
         Caption         =   "C&olors"
         Begin VB.Menu mnuRed 
            Caption         =   "&Red"
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu mnuGreen 
            Caption         =   "&Green"
         End
      End
      Begin VB.Menu mnuSize 
         Caption         =   "&Size"
         Begin VB.Menu mnuBig 
            Caption         =   "&Big"
         End
         Begin VB.Menu mnuSmall 
            Caption         =   "&Small"
         End
      End
   End
End
Attribute VB_Name = "frmShapeColorMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmShapeColorMenu.WindowState = 0
    mnuBig.Checked = False
    mnuSmall.Checked = True
End Sub

Private Sub mnuBig_Click()
    frmShapeColorMenu.WindowState = 2
    mnuBig.Checked = True
    mnuSmall.Checked = False
End Sub

Private Sub mnuBlue_Click()
    Shape1.FillColor = vbBlue
    mnuRed.Checked = False
    mnuBlue.Checked = True
    mnuGreen.Checked = False
End Sub

Private Sub mnuCircle_Click()
    Shape1.Shape = 3
End Sub

Private Sub mnuFileExit_Click()
    Dim x As Integer, str As String
    str = "Are you sure you want to exit?"
    x = MsgBox(str, vbQuestion + vbYesNo, "Exiting")
    If x = vbYes Then
        End
    End If
End Sub

Private Sub mnuFileOpen_Click()
    mnuCut.Enabled = True
    mnuCopy.Enabled = True
End Sub

Private Sub mnuGreen_Click()
    Shape1.FillColor = vbGreen
    mnuRed.Checked = False
    mnuBlue.Checked = False
    mnuGreen.Checked = True
End Sub

Private Sub mnuRectangle_Click()
    Shape1.Shape = 0
End Sub

Private Sub mnuRed_Click()
    Shape1.FillColor = vbRed
    mnuRed.Checked = True
    mnuBlue.Checked = False
    mnuGreen.Checked = False
End Sub

Private Sub mnuSmall_Click()
    frmShapeColorMenu.WindowState = 0
    mnuBig.Checked = False
    mnuSmall.Checked = True
End Sub
