VERSION 5.00
Begin VB.Form Designation_form 
   BorderStyle     =   0  'None
   Caption         =   "Designation Setup"
   ClientHeight    =   7695
   ClientLeft      =   1200
   ClientTop       =   2385
   ClientWidth     =   3135
   Icon            =   "Designation_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   7200
      Width           =   1095
   End
   Begin VB.ListBox lstDesig 
      Height          =   4740
      ItemData        =   "Designation_form.frx":08CA
      Left            =   120
      List            =   "Designation_form.frx":08CC
      TabIndex        =   3
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtDesignationID 
      Height          =   375
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtDesignationDesc 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdSaveDesig 
      Height          =   375
      Left            =   2640
      Picture         =   "Designation_form.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Add entered Record"
      Top             =   960
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   7695
      Left            =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Designation Setting"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3135
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   3000
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   3000
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label1 
      Caption         =   " Designation Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label lblDepartmentID 
      AutoSize        =   -1  'True
      Caption         =   "Designation ID"
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
      TabIndex        =   6
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Designation Description"
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
      TabIndex        =   5
      Top             =   1440
      Width           =   2025
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Desig ID "
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
      Top             =   1800
      Width           =   810
   End
End
Attribute VB_Name = "Designation_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'MsgBox err.Number & " " & err.Description & " cmdSaveDept_Click"


End Sub

Private Sub Command1_Click()
Unload Me
Set rs2 = Nothing
Set rs3 = Nothing
Set rs4 = Nothing
Set rs11 = Nothing
End Sub

Private Sub Command2_Click()
txtDesignationID.Text = ""
txtDesignationDesc.Text = ""
txtDesignationID.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo err
Me.Left = (Screen.Width - Me.Width) / 8
Me.Top = (Screen.Height - Me.Height) / 2
prcConnection
prcAddLstDesignation
rs3.Open "Select * from leaves", con, adOpenDynamic, adLockOptimistic
'rs3.MoveFirst
rs3.Close
rs4.Open "Select * from tax", con, adOpenDynamic, adLockOptimistic
txtpfDedn = Trim(rs4(1))
rs4.Close
Exit Sub
err:
rs3.Close
Exit Sub
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




