VERSION 5.00
Begin VB.Form Department_form 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8175
   ClientLeft      =   795
   ClientTop       =   1365
   ClientWidth     =   3255
   LinkTopic       =   "Form4"
   ScaleHeight     =   8175
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   7680
      Width           =   975
   End
   Begin VB.ListBox lstDepartment 
      Height          =   5520
      ItemData        =   "Department_form.frx":0000
      Left            =   120
      List            =   "Department_form.frx":0002
      TabIndex        =   6
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtDepartmentID 
      Height          =   405
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtDepartmentDesc 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdSaveDept 
      Height          =   375
      Left            =   2760
      Picture         =   "Department_form.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   8295
      Left            =   0
      Top             =   -120
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   3120
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Deptt ID |         Department Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2985
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Department ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Department Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Department Setting"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   -840
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   3120
      Y1              =   7560
      Y2              =   7560
   End
End
Attribute VB_Name = "Department_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSaveDept_Click()
On Error GoTo err
If Len(txtDepartmentID.Text) > 0 And Len(txtDepartmentDesc.Text) > 0 Then
rs1.Open "Select * from dept", con, adOpenDynamic, adLockOptimistic

    rs1.AddNew
    rs1(0) = Trim(txtDepartmentID.Text)
    rs1(1) = Trim(txtDepartmentDesc.Text)
    rs1.Update
    txtDepartmentDesc.Text = ""
    txtDepartmentID.Text = ""
    rs1.Close
        prcAddLstDepartment

Else
    MsgBox "Enter entire details"
End If
Exit Sub
err:
MsgBox err.Number & " " & err.Description & " cmdSaveDept_Click"

End Sub

Private Sub Command1_Click()
txtDepartmentID.Text = ""
txtDepartmentDesc = ""
txtDepartmentID.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 8
Me.Top = (Screen.Height - Me.Height) / 2
prcConnection
prcAddLstDepartment

End Sub

Private Sub lstDepartment_Click()
'On Error GoTo ERR
Dim deptID As String

deptID = Mid(lstDepartment.Text, 1, 2)
rs11.Open "Select * from dept where dept = '" & Trim(deptID) & "'", con, adOpenDynamic, adLockOptimistic
txtDepartmentID.Text = rs11(0)
txtDepartmentDesc.Text = rs11(1)
rs11.Close
Exit Sub
err:
MsgBox err.Number & " " & err.Description & " lstDepartment_Click"
End Sub
Public Sub prcAddLstDepartment()
rs1.Open "Select * from dept", con, adOpenDynamic, adLockOptimistic

rs1.Requery
On Error GoTo err
lstDepartment.Clear
Do While Not rs1.EOF
   lstDepartment.AddItem rs1(0) + "        " + rs1(1)
   rs1.MoveNext
Loop
rs1.Close
Exit Sub
err:
Exit Sub
End Sub
