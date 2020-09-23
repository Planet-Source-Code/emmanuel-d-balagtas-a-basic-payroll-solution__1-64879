VERSION 5.00
Begin VB.Form tax_form 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2775
   ClientLeft      =   990
   ClientTop       =   1950
   ClientWidth     =   4455
   LinkTopic       =   "Form2"
   ScaleHeight     =   2775
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtTaxDedn 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtpfDedn 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "TAX Setting"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Provident Fund Deduction"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Please don't enter % Enter 0.12 if 12% needs to be entered"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Tax Deduction"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1050
   End
End
Attribute VB_Name = "tax_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Not (IsNumeric(Trim(txtTaxDedn.Text)) Or IsNumeric(Trim(txtpfDedn.Text))) Then
    MsgBox "Please enter numeric value only", vbInformation, "Payroll System"
    Exit Sub
End If
If Len(Trim(txtTaxDedn.Text)) > 0 And Len(Trim(txtpfDedn.Text)) > 0 Then
    rs4.Open "Select * from tax", con, adOpenDynamic, adLockOptimistic
    rs4(0) = CDbl(Trim(txtTaxDedn.Text))
    rs4(1) = CDbl(Trim(txtpfDedn.Text))
    rs4.Update
    rs4.Close
Else
    MsgBox "Not Updated", vbInformation, "Payroll System"
End If

End Sub


Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 8
Me.Top = (Screen.Height - Me.Height) / 4
prcConnection
rs4.Open "Select * from tax", con, adOpenDynamic, adLockOptimistic
txtTaxDedn.Text = Trim(rs4(0))
txtpfDedn = Trim(rs4(1))
rs4.Close
Exit Sub


End Sub

Private Sub txtpfDedn_LostFocus()
If Len(Trim(txtpfDedn.Text)) > 0 Then
    txtpfDedn.Text = CDbl(txtpfDedn.Text)
End If
End Sub

Private Sub txtTaxDedn_LostFocus()
If Len(Trim(txtTaxDedn.Text)) > 0 Then
    txtTaxDedn.Text = CDbl(txtTaxDedn.Text)
End If
End Sub
