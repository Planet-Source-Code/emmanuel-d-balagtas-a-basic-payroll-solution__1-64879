VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2100
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
If Len(Trim(Combo1.Text)) > 0 Then
    frmEmpDetails.empNoDialog = Combo1.Text
    MsgBox frmEmpDetails.empNoDialog
    frmEmpDetails.flag = "e"
    frmEmpDetails.prcEnabled
    Unload Me

End If
End Sub

Private Sub Form_Load()
prcConnection
rs.Open "Select * from empmaster", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
    Combo1.AddItem rs("empno")
    rs.MoveNext
Loop
rs.Close
End Sub
