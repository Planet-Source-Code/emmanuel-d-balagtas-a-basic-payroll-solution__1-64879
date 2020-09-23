VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetting 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   12315
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tbSetting 
      Height          =   9135
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   4
      TabHeight       =   520
      TabCaption(0)   =   "Department Master"
      TabPicture(0)   =   "frmSetting.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstDepartment"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDepartmentID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDepartmentDesc"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSaveDept"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Designations Master"
      TabPicture(1)   =   "frmSetting.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "cmdSaveDesig"
      Tab(1).Control(2)=   "txtDesignationDesc"
      Tab(1).Control(3)=   "txtDesignationID"
      Tab(1).Control(4)=   "lstDesig"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(7)=   "lblDepartmentID"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Leave Master"
      TabPicture(2)   =   "frmSetting.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(2)=   "txtWKD"
      Tab(2).Control(3)=   "txtLeavesAvail"
      Tab(2).Control(4)=   "cmdUpdate"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Tax"
      TabPicture(3)   =   "frmSetting.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command2"
      Tab(3).Control(1)=   "txtpfDedn"
      Tab(3).Control(2)=   "txtTaxDedn"
      Tab(3).Control(3)=   "Label10"
      Tab(3).Control(4)=   "Label9"
      Tab(3).Control(5)=   "Label8"
      Tab(3).ControlCount=   6
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   -68640
         TabIndex        =   26
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   -73560
         TabIndex        =   24
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Height          =   375
         Left            =   -73560
         TabIndex        =   23
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox txtpfDedn 
         Height          =   495
         Left            =   -70440
         TabIndex        =   22
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox txtTaxDedn 
         Height          =   495
         Left            =   -70440
         TabIndex        =   20
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtLeavesAvail 
         Height          =   375
         Left            =   -70200
         TabIndex        =   18
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox txtWKD 
         Height          =   375
         Left            =   -70200
         TabIndex        =   16
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton cmdSaveDesig 
         Caption         =   ">>"
         Height          =   495
         Left            =   -67440
         TabIndex        =   14
         ToolTipText     =   "Add entered Record"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtDesignationDesc 
         Height          =   495
         Left            =   -71640
         TabIndex        =   12
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtDesignationID 
         Height          =   495
         Left            =   -71640
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1800
         Width           =   735
      End
      Begin VB.ListBox lstDesig 
         Height          =   5520
         ItemData        =   "frmSetting.frx":0070
         Left            =   -66960
         List            =   "frmSetting.frx":0072
         TabIndex        =   8
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CommandButton cmdSaveDept 
         Caption         =   ">>"
         Height          =   495
         Left            =   7560
         TabIndex        =   7
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtDepartmentDesc 
         Height          =   495
         Left            =   3360
         TabIndex        =   5
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtDepartmentID 
         Height          =   495
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1920
         Width           =   735
      End
      Begin VB.ListBox lstDepartment 
         Height          =   5520
         ItemData        =   "frmSetting.frx":0074
         Left            =   8040
         List            =   "frmSetting.frx":0076
         TabIndex        =   1
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Please don't enter % Enter 0.12 if 12% needs to be entered"
         Height          =   195
         Left            =   -67800
         TabIndex        =   25
         Top             =   3720
         Width           =   4155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Provident Fund Deduction"
         Height          =   195
         Left            =   -73680
         TabIndex        =   21
         Top             =   3480
         Width           =   1860
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tax Deduction"
         Height          =   195
         Left            =   -73680
         TabIndex        =   19
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label7 
         Caption         =   "Leaves Available per month"
         Height          =   375
         Left            =   -73680
         TabIndex        =   17
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Working Days in Organization(Typically 24)"
         Height          =   495
         Left            =   -73680
         TabIndex        =   15
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Desig ID                          Designation Description"
         Height          =   195
         Left            =   -66960
         TabIndex        =   13
         Top             =   1560
         Width           =   3465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Designation Description"
         Height          =   195
         Left            =   -73440
         TabIndex        =   11
         Top             =   2880
         Width           =   1680
      End
      Begin VB.Label lblDepartmentID 
         AutoSize        =   -1  'True
         Caption         =   "Designation ID"
         Height          =   195
         Left            =   -73440
         TabIndex        =   9
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Deptt ID                          Department Description"
         Height          =   195
         Left            =   8040
         TabIndex        =   6
         Top             =   1560
         Width           =   3435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Department Description"
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   2880
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department ID"
         Height          =   195
         Left            =   1560
         TabIndex        =   2
         Top             =   1920
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SSTab1_DblClick()

End Sub

Private Sub cmdDeleteDept_Click()
On Error GoTo err

Dim deptID As String

deptID = Mid(lstDepartment.Text, 1, 2)

rs11.Open "Select * from dept where dept = '" & Trim(deptID) & "'", con, adOpenDynamic, adLockOptimistic
rs11.Delete
rs11.Close
prcAddLstDepartment
Exit Sub
err:
If err.Number = (-2147217887) Then
    MsgBox "Record has been used in another table, So Please delete from the tables where has been used."
End If

End Sub

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

Private Sub cmdSaveDesig_Click()
On Error GoTo err

If Len(txtDesignationID.Text) > 0 And Len(txtDesignationDesc.Text) > 0 Then
    rs2.Open "Select * from desig", con, adOpenDynamic, adLockOptimistic

    rs2.AddNew
    rs2(0) = Trim(txtDesignationID.Text)
    rs2(1) = Trim(txtDesignationDesc.Text)
    rs2.Update
    txtDesignationDesc.Text = ""
    txtDesignationDesc.Text = ""
    rs2.Close
    prcAddLstDesignation
Else
    MsgBox "Enter entire details"
End If
Exit Sub
err:
MsgBox err.Number & " " & err.Description & " cmdSaveDept_Click"

End Sub

Private Sub Command1_Click()
frmAllEmp.Show
End Sub

Private Sub cmdUpdate_Click()
If Not (IsNumeric(Trim(txtWKD.Text)) Or IsNumeric(Trim(txtLeavesAvail.Text))) Then
    MsgBox "Please enter numeric value only"
    Exit Sub
End If

If Len(Trim(txtWKD.Text)) > 0 And Len(Trim(txtLeavesAvail.Text)) > 0 Then
    rs3.Open "Select * from leaves", con, adOpenDynamic, adLockOptimistic
    rs3(0) = Trim(txtWKD.Text)
    rs3(1) = Trim(txtLeavesAvail.Text)
    rs3.Update
    rs3.Close
Else
    MsgBox "Not Updated"
End If
End Sub

Private Sub Command2_Click()
If Not (IsNumeric(Trim(txtTaxDedn.Text)) Or IsNumeric(Trim(txtpfDedn.Text))) Then
    MsgBox "Please enter numeric value only"
    Exit Sub
End If
If Len(Trim(txtTaxDedn.Text)) > 0 And Len(Trim(txtpfDedn.Text)) > 0 Then
    rs4.Open "Select * from tax", con, adOpenDynamic, adLockOptimistic
    rs4(0) = CDbl(Trim(txtTaxDedn.Text))
    rs4(1) = CDbl(Trim(txtpfDedn.Text))
    rs4.Update
    rs4.Close
Else
    MsgBox "Not Updated"
End If

End Sub

Private Sub Form_Load()
On Error GoTo err
prcConnection
prcAddLstDepartment
prcAddLstDesignation
prcDisableMenu
rs3.Open "Select * from leaves", con, adOpenDynamic, adLockOptimistic
'rs3.MoveFirst
txtWKD.Text = Trim(rs3(0))
txtLeavesAvail.Text = Trim(rs3(1))
rs3.Close
rs4.Open "Select * from tax", con, adOpenDynamic, adLockOptimistic
txtTaxDedn.Text = Trim(rs4(0))
txtpfDedn = Trim(rs4(1))
rs4.Close
Exit Sub
err:
rs3.Close
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
prcEnabledMenu
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

Private Sub lstDesig_Click()
On Error GoTo err
Dim desigID As String

desigID = Mid(lstDesig.Text, 1, 2)
rs11.Open "Select * from desig where desig = '" & Trim(desigID) & "'", con, adOpenDynamic, adLockOptimistic
txtDesignationID.Text = rs11(0)
txtDesignationDesc.Text = rs11(1)
rs11.Close
Exit Sub
err:
If err.Number = (3705) Then
    rs11.Close
End If
'MsgBox err.Number & " " & err.Description & " lstDepartment_Click"
End Sub

Private Sub Text3_Change()

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
Public Sub prcAddLstDesignation()
rs2.Open "Select * from desig", con, adOpenDynamic, adLockOptimistic

rs2.Requery
On Error GoTo err
lstDesig.Clear
rs2.MoveFirst
Do While Not rs2.EOF
   lstDesig.AddItem rs2(0) + "        " + rs2(1)
   rs2.MoveNext
Loop
rs2.Close
Exit Sub
err:
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

