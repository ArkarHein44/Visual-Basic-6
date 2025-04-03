VERSION 5.00
Begin VB.Form frmDrawingProgram 
   Caption         =   "Drawing Program"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   480
      Top             =   2280
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPoints 
         Caption         =   "&Points"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu Hyphen1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmDrawingProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Points

Private Sub Form_Load()
    Points = 0
End Sub

Private Sub mnuClear_Click()
    Points = 0
    frmDrawingProgram.Cls
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuPoints_Click()
    Points = 1
End Sub

Private Sub Timer1_Timer()
    Dim R, G, B
    Dim X, Y
    Dim Counter
    If Points = 1 Then
        For Counter = 1 To 100 Step 1
            R = Rnd * 255
            G = Rnd * 255
            B = Rnd * 255
            X = Rnd * frmDrawingProgram.ScaleWidth
            Y = Rnd * frmDrawingProgram.ScaleHeight
            frmDrawingProgram.PSet (X, Y), RGB(R, G, B)
        Next
    End If
End Sub
