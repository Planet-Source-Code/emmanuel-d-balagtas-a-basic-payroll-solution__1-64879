VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEmpDetails 
   Caption         =   "Employee Details"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EmpDetails.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   11100
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   6840
      TabIndex        =   32
      Top             =   8160
      Width           =   2775
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Tag             =   "SANJEEV"
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtPhone 
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Tag             =   "SANJEEV"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.ComboBox cmdDesignation 
      Height          =   360
      ItemData        =   "EmpDetails.frx":1CFA
      Left            =   2520
      List            =   "EmpDetails.frx":1CFC
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   5280
      Width           =   2655
   End
   Begin VB.ComboBox cmdDepartment 
      Height          =   360
      ItemData        =   "EmpDetails.frx":1CFE
      Left            =   2520
      List            =   "EmpDetails.frx":1D00
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4560
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   240
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      Picture         =   "EmpDetails.frx":1D02
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      Picture         =   "EmpDetails.frx":25CC
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      Picture         =   "EmpDetails.frx":2E96
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Picture         =   "EmpDetails.frx":3760
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox txtLname 
      Height          =   360
      Left            =   2520
      TabIndex        =   0
      Tag             =   "SANJEEV"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Browse"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   7440
      Width           =   2175
   End
   Begin VB.PictureBox pcEmpPicture 
      DataField       =   "Picture"
      DataSource      =   "Adodc1"
      Height          =   2415
      Left            =   5520
      ScaleHeight     =   2355
      ScaleWidth      =   2115
      TabIndex        =   22
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtPFD 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   7440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtBase 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox txtPFN 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox txtDOJ 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox txtAddress 
      Height          =   1095
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox txtFname 
      Height          =   360
      Left            =   6720
      TabIndex        =   1
      Tag             =   "SANJEEV"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   0
      Picture         =   "EmpDetails.frx":545A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15465
   End
   Begin VB.Label Label3 
      Caption         =   "Please select employee ID"
      Height          =   255
      Left            =   6840
      TabIndex        =   33
      Top             =   8520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "DD / MM / YYYY"
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email"
      Height          =   240
      Left            =   7080
      TabIndex        =   30
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblPhone 
      AutoSize        =   -1  'True
      Caption         =   "Phone"
      Height          =   240
      Left            =   6960
      TabIndex        =   29
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label txtEmpPicturePath 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   5520
      TabIndex        =   28
      Top             =   6960
      Width           =   60
   End
   Begin VB.Label lblPicture 
      Caption         =   "Employee's Photo"
      Height          =   255
      Left            =   5760
      TabIndex        =   23
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblPFD 
      AutoSize        =   -1  'True
      Caption         =   "Provident Fund Dedn"
      Height          =   240
      Left            =   480
      TabIndex        =   21
      Top             =   7560
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblBase 
      AutoSize        =   -1  'True
      Caption         =   "Basic Pay"
      Height          =   240
      Left            =   480
      TabIndex        =   20
      Top             =   6840
      Width           =   900
   End
   Begin VB.Label lblPFN 
      AutoSize        =   -1  'True
      Caption         =   "Provident Fund No"
      Height          =   240
      Left            =   480
      TabIndex        =   19
      Top             =   6000
      Width           =   1590
   End
   Begin VB.Label lblDesig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      Height          =   240
      Left            =   480
      TabIndex        =   18
      Top             =   5280
      Width           =   1020
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   240
      Left            =   480
      TabIndex        =   17
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Label lblDOJ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Hired"
      Height          =   240
      Left            =   480
      TabIndex        =   16
      Top             =   3840
      Width           =   915
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   240
      Left            =   480
      TabIndex        =   15
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lblFirstName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   240
      Left            =   5640
      TabIndex        =   14
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label lblLastName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   240
      Left            =   480
      TabIndex        =   13
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   7800
      TabIndex        =   12
      Top             =   3480
      Width           =   3015
   End
End
Attribute VB_Name = "frmEmpDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public empNoDialog As String

Dim control As Variant
Dim errorFlag As Boolean
Public flag As String
Dim newEmp As String
Dim empName As String

Private Sub cmdCancel_Click()
On Error GoTo ERRR
rs.CancelUpdate
rs.Close
prcDisabled
flag = ""
Exit Sub
ERRR:
'rs.Close
prcDisabled
flag = ""
Exit Sub
End Sub

Private Sub cmdEdit_Click()
flag = "e"
cmdEdit.Visible = False
Label3.Visible = True
prcEnabled
Combo1.Visible = True
Combo1.Top = cmdEdit.Top

End Sub

Private Sub cmdNew_Click()
On Error GoTo err
prcNewEmp
prcEnabled
rs.Open "Select * from empmaster", con, adOpenDynamic, adLockOptimistic
rs.AddNew
flag = "n"
Exit Sub
err:
MsgBox "eRROR"
End Sub

Private Sub cmdSave_Click()
prcValidate
If errorFlag = True Then
    flag = ""
    prcSave
    rs.Update
    rs.Close
        MsgBox "Employee Id for " & txtFname.Text & " " & txtLname.Text & " is " & newEmp

    prcDisabled
    
End If
End Sub

Private Sub Combo1_Click()
Combo1.Clear
rs.Open "SELECT * FROM EMPMASTER", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
Combo1.AddItem rs("EMPNO")
rs.MoveNext
Loop
rs.Close

rs.Open "SELECT * FROM EMPMASTER WHERE EMPNO ='" & Trim(Combo1.Text) & "'", con, adOpenDynamic, adLockOptimistic
newEmp = Combo1.Text
cmdEdit.Visible = True
Combo1.Visible = False
Label3.Visible = False
prcOpenRecord
End Sub

Private Sub Command4_Click()
CD1.DialogTitle = "Open the Employee's Picture"
CD1.InitDir = "C:\"
CD1.Filter = "All File(*.*)|*.*|JPEG Files(*.jpg)|*.jpg|BMP Files(*.bmp)|*.bmp"
CD1.ShowOpen
txtEmpPicturePath.Caption = CD1.FileName
pcEmpPicture.Picture = LoadPicture(txtEmpPicturePath.Caption)
pcEmpPicture.Visible = True
End Sub
Private Sub Form_Load()

prcConnection
errorFlag = True
cmdDepartment.AddItem " "
cmdDesignation.AddItem " "
rs.Open "Select * from dept", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
cmdDepartment.AddItem rs(0)
rs.MoveNext
Loop
rs.Close

rs.Open "Select * from desig", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
cmdDesignation.AddItem rs(0)
rs.MoveNext
Loop
rs.Close
rs.Open "SELECT * FROM EMPMASTER", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
Combo1.AddItem rs("EMPNO")
rs.MoveNext
Loop
rs.Close
prcDisabled
Combo1.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Len(Trim(flag)) > 0 Then
    MsgBox "you can't exit at this time"
    Cancel = 1
    Exit Sub
End If


End Sub
Public Sub prcDisabled()
For Each control In frmEmpDetails
    If TypeOf control Is TextBox Or TypeOf control Is ComboBox Or TypeOf control Is CommandButton Then
        control.Enabled = False
    End If
Next
cmdNew.Enabled = True
cmdEdit.Enabled = True
Combo1.Visible = False
cmdEdit.Visible = True
prcClear
End Sub

Public Sub prcEnabled()
For Each control In frmEmpDetails
    If TypeOf control Is TextBox Or TypeOf control Is ComboBox Or TypeOf control Is CommandButton Then
        control.Enabled = True
    End If
Next
cmdNew.Enabled = False
cmdEdit.Enabled = False
txtLname.SetFocus
End Sub


Public Sub prcValidate()
Dim i, a As Integer
Dim ch As String
Dim chAT As Boolean
Dim chDOT As Boolean
Dim poAT As Integer
Dim poDOT As Integer
If Len(Trim(txtLname.Text)) = 0 Then
    MsgBox "Please enter Last name of employee"
    errorFlag = False
    lblLastName.ForeColor = vbRed
    txtLname.SetFocus
    Exit Sub
ElseIf IsNumeric(txtLname.Text) Then
    MsgBox "Invalid entry in last name"
    errorFlag = False
    lblLastName.ForeColor = vbRed
    txtLname.SetFocus
    Exit Sub
Else
    errorFlag = True
    lblLastName.ForeColor = vbBlack
End If

If Len(Trim(txtFname.Text)) = 0 Then
    MsgBox "Please enter first name of employee"
    errorFlag = False
    lblFirstName.ForeColor = vbRed
    txtFname.SetFocus
    Exit Sub
ElseIf IsNumeric(txtFname.Text) Then
    MsgBox "Invalid entry in first name"
    errorFlag = False
    lblFirstName.ForeColor = vbRed
    txtFname.SetFocus
    Exit Sub
Else
    errorFlag = True
    lblFirstName.ForeColor = vbBlack

End If

If Len(Trim(txtAddress.Text)) = 0 Then
    MsgBox "Please enter address of employee"
    errorFlag = False
    lblAddress.ForeColor = vbRed
    txtAddress.SetFocus
    Exit Sub
Else
    errorFlag = True
    lblAddress.ForeColor = vbBlack
End If

If Len(Trim(txtDOJ.Text)) = 0 Then
    MsgBox "Please enter Date of Joining of employee"
    errorFlag = False
    lblDOJ.ForeColor = vbRed
    txtDOJ.SetFocus
    Exit Sub
ElseIf Not IsDate(txtDOJ.Text) Then
    MsgBox "Invalid entry in Date of Joining"
    errorFlag = False
    lblDOJ.ForeColor = vbRed
    txtDOJ.SetFocus
    Exit Sub
Else
    errorFlag = True
    lblDOJ.ForeColor = vbBlack
End If

If Len(Trim(txtPhone.Text)) = 0 Then
    MsgBox "Please enter Phone number of employee"
    errorFlag = False
    lblPhone.ForeColor = vbRed
    txtPhone.SetFocus
    Exit Sub

Else
    errorFlag = True
    lblPhone.ForeColor = vbBlack
End If

For i = 1 To Len(Trim(txtEmail.Text))
ch = Mid(Trim(txtEmail.Text), i, Len(Trim(txtEmail.Text)) - (Len(Trim(txtEmail.Text)) - 1))
If ch = "@" Then
    If chAT = True Then
        MsgBox "Wrong Email ID"
        errorFlag = False
        txtEmail.SetFocus
        lblEmail.ForeColor = vbRed
    
        Exit Sub
    Else
        errorFlag = True
        lblEmail.ForeColor = vbBlack
    End If
    
    chAT = True
    poAT = i
ElseIf ch = "." Then
    chDOT = True
    poDOT = i
End If

Next i
If chAT = False Then
    MsgBox "Wrong Email ID missing:" & vbNewLine & "'@'"
    
    errorFlag = False
    txtEmail.SetFocus
    lblEmail.ForeColor = vbRed
    Exit Sub
Else
    errorFlag = True
    lblEmail.ForeColor = vbBlack
End If

If chDOT = False Then
    MsgBox "Wrong Email ID missing:" & vbNewLine & "'.'"
    
    errorFlag = False
    txtEmail.SetFocus
    lblEmail.ForeColor = vbRed
    Exit Sub
Else
    errorFlag = True
    lblEmail.ForeColor = vbBlack
End If

If poAT > poDOT Then
    MsgBox "Wrong Email ID"
    errorFlag = False
    txtEmail.SetFocus
    lblEmail.ForeColor = vbRed
    Exit Sub
Else
    errorFlag = True
    lblEmail.ForeColor = vbBlack
End If

If Len(Trim(cmdDepartment.Text)) = 0 Then
    MsgBox "Select appropriate value for Department"
    errorFlag = False
    cmdDepartment.SetFocus
    lblDept.ForeColor = vbRed
    Exit Sub
Else
    errorFlag = True
    lblDept.ForeColor = vbBlack
End If
If Len(Trim(cmdDesignation.Text)) = 0 Then
    MsgBox "Select appropriate value for Designation"
    errorFlag = False
    cmdDesignation.SetFocus
    lblDesig.ForeColor = vbRed
    Exit Sub
Else
    errorFlag = True
    lblDesig.ForeColor = vbBlack
End If

If Len(Trim(txtPFN.Text)) = 0 Then
    MsgBox "Please enter Provident Fund number of employee"
    errorFlag = False
    lblPFN.ForeColor = vbRed
    txtPFN.SetFocus
    Exit Sub

Else
    errorFlag = True
    lblPFN.ForeColor = vbBlack
End If

If Len(Trim(txtBase.Text)) = 0 Then
    MsgBox "Please enter Basic of employee"
    errorFlag = False
    lblBase.ForeColor = vbRed
    txtBase.SetFocus
    Exit Sub

Else
    errorFlag = True
    lblBase.ForeColor = vbBlack
End If

If Len(Trim(txtPFD.Text)) = 0 Then
    txtPFD.Text = "0"
    errorFlag = False
    lblPFD.ForeColor = vbRed
    'txtPFD.SetFocus
    Exit Sub

Else
    errorFlag = True
    lblPFD.ForeColor = vbBlack
End If
If Len(Trim(txtEmpPicturePath.Caption)) = 0 Then
    If vbNo = MsgBox("Are you sure you don't want to add picture of the employee", vbYesNo, "Confirm Picture Select") Then
        MsgBox "Click on the browse to select the picture"
        errorFlag = False
        Command4.SetFocus
        lblPicture.ForeColor = vbRed
    Else
        errorFlag = True
        lblPicture.ForeColor = vbBlack
    End If
End If
End Sub

Private Sub txtDOJ_LostFocus()
txtDOJ.Text = Format(txtDOJ.Text, "DD/MMM/YYYY")

End Sub

Public Sub prcSave()
On Error GoTo err

rs("EmpNo") = newEmp
rs("First_Name") = Trim(txtFname.Text)
rs("Last_name") = Trim(txtLname.Text)
rs("address") = Trim(txtAddress.Text)
rs("doj") = Trim(txtDOJ.Text)
rs("dept") = Trim(cmdDepartment.Text)
rs("desig") = Trim(cmdDesignation.Text)
rs("pfno") = Trim(txtPFN.Text)
rs("basic") = Trim(txtBase.Text)
rs("pfdedn") = 0
rs("picture_path") = Trim(txtEmpPicturePath.Caption)
rs("phone") = Trim(txtPhone.Text)
rs("email") = Trim(txtEmail.Text)
Exit Sub
err:
MsgBox err.Description
End Sub


Public Sub prcNewEmp()
Dim C As Integer
Dim empNo As String

newEmp = ""
C = 0
rs.Open "Select * from empmaster", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
C = C + 1
rs.MoveNext
Loop
rs.Close

If C = 0 Then
newEmp = "E001"
Exit Sub
End If
rs.Open "Select * from empmaster", con, adOpenDynamic, adLockOptimistic
rs.MoveLast
empNo = rs("empno")
empNo = Mid(empNo, 2, 4)
newEmp = empNo + 1
If newEmp >= 1 And newEmp <= 9 Then
newEmp = "E00" + newEmp
ElseIf newEmp >= 10 And newEmp <= 99 Then
newEmp = "E0" + newEmp
ElseIf newEmp >= 100 And newEmp <= 999 Then
newEmp = "E" + newEmp
End If
rs.Close
End Sub

Public Sub prcClear()
For Each control In frmEmpDetails
    If TypeOf control Is TextBox Then
        control.Text = ""
    End If
    
Next

cmdDepartment.ListIndex = 0
cmdDesignation.ListIndex = 0
Combo1.Text = ""

txtEmpPicturePath.Caption = ""
pcEmpPicture.Visible = False

End Sub

Public Sub prcOpenRecord()
txtFname.Text = rs("First_Name")
txtLname.Text = rs("Last_name")
txtAddress.Text = rs("address")
txtDOJ.Text = rs("doj")
cmdDepartment.Text = rs("dept")
cmdDesignation.Text = rs("desig")
txtPFN.Text = rs("pfno")
txtBase.Text = rs("basic")
txtPFD.Text = rs("pfdedn")
txtEmpPicturePath.Caption = rs("picture_path") & " "
If Len(Trim(txtEmpPicturePath.Caption)) > 0 Then
    pcEmpPicture.Picture = LoadPicture(Trim(txtEmpPicturePath.Caption))
    pcEmpPicture.Visible = True
End If
txtPhone.Text = rs("phone")
txtEmail.Text = rs("email")

End Sub
