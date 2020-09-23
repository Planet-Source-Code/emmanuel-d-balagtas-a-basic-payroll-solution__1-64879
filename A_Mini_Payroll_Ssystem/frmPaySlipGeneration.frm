VERSION 5.00
Begin VB.Form frmPaySlipGeneration 
   Caption         =   "Pay Slip"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPaySlipGeneration.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   10200
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
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
      Height          =   375
      Left            =   240
      TabIndex        =   36
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   2520
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   2040
      TabIndex        =   22
      Top             =   6120
      Width           =   12135
      Begin VB.CommandButton Command2 
         Caption         =   "&Save and Print"
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
         Left            =   840
         Picture         =   "frmPaySlipGeneration.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Line Line9 
         X1              =   5160
         X2              =   6600
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line8 
         X1              =   5160
         X2              =   6600
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line7 
         X1              =   5160
         X2              =   6600
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label txtNetSalary 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5160
         TabIndex        =   34
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Salary --------------------------------------------"
         Height          =   270
         Left            =   240
         TabIndex        =   33
         Top             =   3240
         Width           =   4440
      End
      Begin VB.Label txtAddAllowance 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5160
         TabIndex        =   32
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add: Allowances"
         Height          =   270
         Left            =   720
         TabIndex        =   31
         Top             =   2400
         Width           =   1740
      End
      Begin VB.Label txtTaxDedn 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5160
         TabIndex        =   30
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less:Tax Deduction"
         Height          =   270
         Left            =   720
         TabIndex        =   29
         Top             =   1920
         Width           =   2070
      End
      Begin VB.Label txtPFDeduction 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5160
         TabIndex        =   28
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less: PF Deduction"
         Height          =   270
         Left            =   720
         TabIndex        =   27
         Top             =   1440
         Width           =   2070
      End
      Begin VB.Label txtLeaveDedn 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5160
         TabIndex        =   26
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less: Leave Deduction"
         Height          =   270
         Left            =   720
         TabIndex        =   25
         Top             =   960
         Width           =   2400
      End
      Begin VB.Label txtBaseSalary 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5160
         TabIndex        =   24
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Salary -----------------------------------------"
         Height          =   270
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   4410
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process PaySlip"
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
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtPFDed 
      Enabled         =   0   'False
      Height          =   390
      Left            =   10680
      TabIndex        =   20
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtTaxDeduction 
      Enabled         =   0   'False
      Height          =   390
      Left            =   10680
      TabIndex        =   18
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtLeavesAvailable 
      Enabled         =   0   'False
      Height          =   390
      Left            =   10680
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtTotWorkDays 
      Enabled         =   0   'False
      Height          =   390
      Left            =   10680
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtTotalLeave 
      Height          =   390
      Left            =   3240
      TabIndex        =   9
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtAllowDesc 
      Height          =   975
      Left            =   3240
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox txtAllowance 
      Height          =   390
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ComboBox cmdEmpNo 
      Height          =   390
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   0
      Picture         =   "frmPaySlipGeneration.frx":74DC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15465
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   15480
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provident Fund Deduction"
      Height          =   270
      Left            =   6600
      TabIndex        =   19
      Top             =   4800
      Width           =   2700
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Deduction"
      Height          =   270
      Left            =   6600
      TabIndex        =   17
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leaves Available"
      Height          =   270
      Left            =   6600
      TabIndex        =   15
      Top             =   3600
      Width           =   1785
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Working Days"
      Height          =   270
      Left            =   6600
      TabIndex        =   13
      Top             =   3000
      Width           =   2040
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8760
      TabIndex        =   12
      Top             =   2520
      Width           =   945
   End
   Begin VB.Line Line4 
      X1              =   6480
      X2              =   12120
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line3 
      X1              =   12120
      X2              =   12120
      Y1              =   2400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   12120
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      X1              =   6480
      X2              =   6480
      Y1              =   2400
      Y2              =   5400
   End
   Begin VB.Label txtLeavesWithoutPay 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave without Pay"
      Height          =   270
      Left            =   240
      TabIndex        =   10
      Top             =   4680
      Width           =   1890
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Leave in Month"
      Height          =   270
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   2145
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
      Caption         =   "Description of Allowance"
      Height          =   270
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   2580
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allowance If Any"
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1680
   End
   Begin VB.Label txtMonth 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      TabIndex        =   3
      Top             =   1920
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " For the Month"
      Height          =   270
      Left            =   6840
      TabIndex        =   2
      Top             =   1440
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emp No"
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "frmPaySlipGeneration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As String
Dim payslipid As String

Dim empNos As String


Private Sub Command1_Click()
Dim bal As Integer
If Len(Trim(cmdEmpNo.Text)) = 0 Then
    MsgBox "Please enter Employee No"
    Exit Sub
End If

txtAllowance_LostFocus

txtTotalLeave_LostFocus
bal = CInt(txtTotalLeave.Text) - CInt(txtLeavesAvailable.Text)
If bal >= 0 Then
    txtLeavesWithoutPay.Caption = bal
Else
    txtLeavesWithoutPay.Caption = "0"
End If
TimerOn (2)

Command1.Enabled = False
prcCalculateBasicSalary
flag = "e"
End Sub

Private Sub Command2_Click()
On Error GoTo err
rs.Open "select * from pay_trans", con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs(0) = Trim(empNos)
rs(1) = Trim(CDbl(txtBaseSalary.Caption))
rs(2) = Trim(CDbl(txtTaxDedn.Caption))
rs(3) = Trim(CDbl(txtPFDeduction.Caption))
rs(4) = Trim(CDbl(txtAddAllowance.Caption))
rs(5) = Trim((txtAllowDesc.Text))
rs(6) = Trim(CDbl(txtLeavesWithoutPay.Caption))
rs(7) = Trim(CDbl(txtTotalLeave.Text))
rs(8) = Trim(CDbl(txtLeaveDedn.Caption))
rs(9) = Trim(CDbl(txtNetSalary.Caption))
rs(10) = Trim((txtMonth.Caption))
payslipid = Format(txtMonth.Caption, "MMMYY") & Trim(empNos)
rs("PAYSLIPID") = payslipid
rs.Update
rs.Close
MsgBox "Payslip has been generated, In order to print the PaySlip got Reprint Payslip Menu." & vbNewLine & " Pay Slip ID is " & UCase(payslipid)
flag = ""
Unload Me



Exit Sub
err:
MsgBox "Error occured" & vbNewLine & "Most Probably Pay for this employee has already been generated", vbExclamation, "Payroll Solution"
flag = ""
Unload Me
End Sub



Private Sub Command3_Click()
Unload Me
Set rs = Nothing
End Sub

Private Sub Form_Load()

Frame1.Top = 12000
prcConnection
rs.Open "Select * from empmaster", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
    cmdEmpNo.AddItem rs(0) & "--" & rs(2) & " " & rs(1)
    rs.MoveNext
Loop
rs.Close
rs.Open "Select * from leaves", con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
txtTotWorkDays = Trim(rs(0))
txtLeavesAvailable = Trim(rs(1))
rs.Close

rs.Open "Select * from tax", con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
txtTaxDeduction = Trim(rs(0))
txtPFDed = Trim(rs(1))
rs.Close
txtMonth.Caption = Format(Now, "MMM/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(flag) > 0 Then
    If vbNo = MsgBox("Are you sure you want to exit without saving", vbCritical + vbYesNo) Then
    Cancel = 1
    
    Exit Sub
    End If
End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()
If Frame1.Top > 2000 Then
    Frame1.Top = Frame1.Top - 200
Else
    TimerOn (0)
End If

End Sub

Private Sub txtAllowance_LostFocus()
If Len(Trim(txtAllowance.Text)) > 0 Then
If Not IsNumeric(txtAllowance.Text) Then
    MsgBox "Enter numeric value only"
End If
End If
If Len(Trim(txtAllowance.Text)) = 0 Then
    txtAllowance.Text = "0"
End If

End Sub

Private Sub txtTotalLeave_LostFocus()
If Len(Trim(txtTotalLeave.Text)) > 0 Then
If Not IsNumeric(txtTotalLeave.Text) Then
    MsgBox "Enter numeric value only"
End If
End If
If Len(Trim(txtTotalLeave.Text)) = 0 Then
    txtTotalLeave.Text = "0"
End If

End Sub
Public Function TimerOn(Interval As Integer)

   If Interval > 0 Then
      Timer1.Interval = Interval   ' Enables the timer.
   Else
      Timer1.Interval = 0         ' Disables the timer.
   End If
End Function

Public Sub prcCalculateBasicSalary()
Dim netSal As Double
Dim onedaySal As Double
empNos = Mid(Trim(cmdEmpNo.Text), 1, 4)
rs.Open "Select * from empmaster where empno='" & empNos & "'", con, adOpenDynamic, adLockOptimistic
txtBaseSalary.Caption = rs("basic")
rs.Close
txtPFDeduction.Caption = CDbl(txtBaseSalary.Caption) * CDbl(txtPFDed.Text)
txtTaxDedn.Caption = CDbl(txtBaseSalary.Caption) * CDbl(txtTaxDeduction.Text)
txtAddAllowance.Caption = CDbl(txtAllowance.Text)
onedaySal = CDbl(txtBaseSalary.Caption) / CDbl(txtTotWorkDays.Text)
txtLeaveDedn = onedaySal * CDbl(txtLeavesWithoutPay.Caption)
netSal = CDbl(txtBaseSalary) - (CDbl(txtLeaveDedn) + CDbl(txtPFDeduction) + CDbl(txtTaxDedn)) + CDbl(txtAddAllowance)
txtNetSalary = netSal
txtNetSalary = CInt(netSal)
End Sub
