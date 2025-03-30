VERSION 5.00
Begin VB.Form frmLoaction2initialPoints 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Location Two Initial Points"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
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
      Left            =   1553
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "frmLoaction2initialPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Example (2.1.9)  Given f(x) = x2 -10x - 14 = 0. Write a Visual Basic Program
'for locating two initial points x1 and x2  such that f(x1) and f(x2) have different signs
'(so as ready for Bisection Method).

Private Function f(x As Double) As Double
    f = x ^ 2 - 10 * x - 14
End Function

Private Sub cmdStart_Click()
    ' f(x) = x^2 -10x - 14
    
    Dim a As Double, b As Double
    Dim i As Integer, j As Integer
    
    For i = -6 To 6
        j = i + 1
        a = f(CDbl(i))
        b = f(CDbl(j))
        
        ' Display current values being checked
        MsgBox "x1 = " & CStr(i) & ", y1 = " & CStr(a) & vbNewLine & _
               "x2 = " & CStr(j) & ", y2 = " & CStr(b)
        
        If a * b < 0 Then
            MsgBox "There is a solution between x1 = " & CStr(i) & ", x2 = " & CStr(j) & vbNewLine & _
                   "with y1 = " & CStr(a) & ", y2 = " & CStr(b)
            txtResult.Text = "There is a solution between " & CStr(i) & " and " & CStr(j)
            Exit For ' This stops the looping if a solution is found
        End If
    Next i
    
    If a * b >= 0 Then
        txtResult.Text = "No solution found in the specified range"
    End If
End Sub

Private Sub Text1_Change()

End Sub
