VERSION 5.00
Begin VB.Form Leave_form 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2415
   ClientLeft      =   990
   ClientTop       =   2145
   ClientWidth     =   3975
   LinkTopic       =   "Form2"
   ScaleHeight     =   2415
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtWKD 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtLeavesAvail 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   2415
      Left            =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Working Days in Organization(Typically 24)"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Leaves Available per month"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Leave Setting"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   -360
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Leave_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUpdate_Click()
If Not (IsNumeric(Trim(txtWKD.Text)) Or IsNumeric(Trim(txtLeavesAvail.Text))) Then
    MsgBox "Please enter numeric value only", vbInformation, "Mini Payroll System"
    
    Exit Sub
End If

If Len(Trim(txtWKD.Text)) > 0 And Len(Trim(txtLeavesAvail.Text)) > 0 Then
    rs3.Open "Select * from leaves", con, adOpenDynamic, adLockOptimistic
    rs3(0) = Trim(txtWKD.Text)
    rs3(1) = Trim(txtLeavesAvail.Text)
    rs3.Update
    rs3.Close
Else
    MsgBox "Not Updated", vbInformation, "Mini Payroll System"
End If
End Sub

Private Sub Command1_Click()
Unload Me
Set rs3 = Nothing
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 8
Me.Top = (Screen.Height - Me.Height) / 4
prcConnection
rs3.Open "Select * from leaves", con, adOpenDynamic, adLockOptimistic
'rs3.MoveFirst
txtWKD.Text = Trim(rs3(0))
txtLeavesAvail.Text = Trim(rs3(1))

End Sub
