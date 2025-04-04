VERSION 5.00
Begin VB.Form frmChangeToDay 
   Caption         =   "Change to Day"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
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
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
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
      Height          =   495
      Left            =   2333
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblOutput 
      Alignment       =   2  'Center
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
      Left            =   173
      TabIndex        =   2
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "No of Day = "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   893
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmChangeToDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
    If IsNumeric(Text1.Text) Then
        Select Case CInt(Text1.Text)
            Case 1
                lblOutput.Caption = "Your selected day is Sunday"
                lblOutput.BackColor = &HC0FFC0
                lblOutput.ForeColor = vbBlack
                MsgBox "Your selected day is Sunday", vbInformation + vbOKOnly, "Name of Day"
            Case 2
                lblOutput.Caption = "Your selected day is Monday"
                lblOutput.BackColor = &HC0FFC0
                lblOutput.ForeColor = vbBlack
                MsgBox "Your selected day is Monday", vbInformation + vbOKOnly, "Name of Day"
            Case 3
                lblOutput.Caption = "Your selected day is Tuesday"
                lblOutput.BackColor = &HC0FFC0
                lblOutput.ForeColor = vbBlack
                MsgBox "Your selected day is Tuesday", vbInformation + vbOKOnly, "Name of Day"
            Case 4
                lblOutput.Caption = "Your selected day is Wednesday"
                lblOutput.BackColor = &HC0FFC0
                lblOutput.ForeColor = vbBlack
                MsgBox "Your selected day is Wednesday", vbInformation + vbOKOnly, "Name of Day"
            Case 5
                lblOutput.Caption = "Your selected day is Thursday"
                lblOutput.BackColor = &HC0FFC0
                lblOutput.ForeColor = vbBlack
                MsgBox "Your selected day is Thursday", vbInformation + vbOKOnly, "Name of Day"
            Case 6
                lblOutput.Caption = "Your selected day is Friday"
                lblOutput.BackColor = &HC0FFC0
                lblOutput.ForeColor = vbBlack
                MsgBox "Your selected day is Friday", vbInformation + vbOKOnly, "Name of Day"
            Case 7
                lblOutput.Caption = "Your selected day is Saturday"
                lblOutput.BackColor = &HC0FFC0
                lblOutput.ForeColor = vbBlack
                MsgBox "Your selected day is Saturday", vbInformation + vbOKOnly, "Name of Day"
            Case Else
                lblOutput.Caption = "day must be between 1 to 7"
                lblOutput.ForeColor = vbRed
                lblOutput.BackColor = &H8000000F
                Text1.Text = ""
                Text1.SetFocus
        End Select
    Else
        MsgBox "Only number Allow!!!", vbExclamation + vbOKOnly, "Invalid input"
        Text1.SetFocus
    End If
End Sub
