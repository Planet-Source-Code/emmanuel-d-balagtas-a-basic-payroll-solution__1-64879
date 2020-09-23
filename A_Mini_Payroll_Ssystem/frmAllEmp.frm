VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAllEmp 
   Caption         =   "All Details"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   10680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Height          =   735
      Left            =   2280
      Picture         =   "frmAllEmp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   1200
      Picture         =   "frmAllEmp.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      Picture         =   "frmAllEmp.frx":39F4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmAllEmp.frx":56EE
      Height          =   7440
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   13123
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      DataMember      =   "comALL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
      _Band(0)._NumMapCols=   11
      _Band(0)._MapCol(0)._Name=   "Employee_ID"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "Name"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "Date_of_Joining"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "Department"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "Designation"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "MonthOfPay"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "Basic"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(6)._Alignment=   7
      _Band(0)._MapCol(7)._Name=   "PF_Deduction"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(7)._Alignment=   7
      _Band(0)._MapCol(8)._Name=   "LeaveDeduction"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(8)._Alignment=   7
      _Band(0)._MapCol(9)._Name=   "Tax_Deduction"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(9)._Alignment=   7
      _Band(0)._MapCol(10)._Name=   "Net_Pay"
      _Band(0)._MapCol(10)._RSIndex=   10
      _Band(0)._MapCol(10)._Alignment=   7
   End
   Begin VB.Label Label3 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Delete"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Print"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   8400
      Width           =   495
   End
End
Attribute VB_Name = "frmAllEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
rptPaySlip.Show
Unload Me
End Sub

Private Sub Command2_Click()
MsgBox "DEMO VERSION", vbInformation
End Sub

Private Sub Command3_Click()
Unload Me

End Sub



Private Sub Form_Load()
With MSHFlexGrid1
    .ColWidth(0) = 1500
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    .ColWidth(9) = 1500
    .ColWidth(10) = 1500
End With
  
End Sub

Private Sub Form_Resize()
'MSHFlexGrid1.Height = Me.Height - 400
'MSHFlexGrid1.Width = Me.Width - 20

End Sub



