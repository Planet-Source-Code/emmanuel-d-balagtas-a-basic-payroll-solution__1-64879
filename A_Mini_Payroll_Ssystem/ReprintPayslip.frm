VERSION 5.00
Begin VB.Form frmReprintPayslip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "ReprintPayslip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   8655
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   3135
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton frmReprintPayslip 
      Caption         =   "Reprint Pay Slip"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "To Date"
      Height          =   195
      Left            =   4440
      TabIndex        =   4
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From Date"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee ID"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frmReprintPayslip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Call prcCallAllEmp(CStr(Combo2.Text), CStr(Combo3.Text))

Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
Set rs = Nothing
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 8
Me.Top = (Screen.Height - Me.Height) / 4
If typeID = "c" Then
    frmReprintPayslip.Enabled = False
    Combo1.Enabled = False
End If
prcConnection

rs.Open "select * from PaySlipGenerated", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
If typeID = "a" Then
Combo1.AddItem rs("payslipid")
ElseIf typeID = "b" Then
Combo1.AddItem rs("empno") + " " + rs("name")
End If
rs.MoveNext
Loop
rs.Close
rs.Open "Select distinct(month) from payslipgenerated", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
Combo2.AddItem Format(rs(0), "mmm-yyyy")
Combo3.AddItem Format(rs(0), "mmm-yyyy")
rs.MoveNext
Loop
rs.Close

End Sub





Private Sub frmReprintPayslip_Click()
If typeID = "a" Then
prcCallPaySlip (Combo1.Text)
ElseIf typeID = "b" Then
prcCallempPaySlip (Mid(Combo1.Text, 1, 4))
End If
typeID = ""
Unload Me
End Sub
