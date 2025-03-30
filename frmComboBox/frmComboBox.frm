VERSION 5.00
Begin VB.Form frmComboBox 
   Caption         =   "Combo Box"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboUniv 
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
      ItemData        =   "frmComboBox.frx":0000
      Left            =   1013
      List            =   "frmComboBox.frx":0002
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtIntake 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1013
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    cboUniv.AddItem "Yangon University"
    cboUniv.AddItem "Mandalay University"
    cboUniv.AddItem "Pathein University"
    cboUniv.AddItem "Sittway University"
End Sub

Private Sub cboUniv_Click()
    txtIntake.Text = cboUniv.Text
End Sub

