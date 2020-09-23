VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMainForm 
   BackColor       =   &H8000000C&
   Caption         =   "EDB Payroll Solution"
   ClientHeight    =   7125
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11010
   Icon            =   "frmMainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6870
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Picture         =   "frmMainForm.frx":08CA
            TextSave        =   "12:43 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "frmMainForm.frx":0A69
            TextSave        =   "3/13/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11360
            MinWidth        =   11360
            Picture         =   "frmMainForm.frx":0CDC
            Text            =   "Copyright 2006 Emmanuel D. Balagtas | www.edbsoft.cjb.net | edbsoft@gmail.com"
            TextSave        =   "Copyright 2006 Emmanuel D. Balagtas | www.edbsoft.cjb.net | edbsoft@gmail.com"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employee Information"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Payroll Settings"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pay Slip Generator"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Pay Slip"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   18
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":0E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":2611
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":431B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":5AAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":77B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":8F49
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":A6DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":C3E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":E0EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":FDF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":1158B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":13295
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":14A27
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":16731
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":17EC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":18B9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":1A32F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":1C039
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":1DD43
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuEmployee 
         Caption         =   "Employee"
         Shortcut        =   ^E
      End
      Begin VB.Menu bar01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuadmin 
         Caption         =   "&Administrator"
         Shortcut        =   ^A
      End
      Begin VB.Menu bar02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuSet 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSetDesig 
         Caption         =   "Designation Setting"
      End
      Begin VB.Menu mnuLeave 
         Caption         =   "&Leav Setting"
      End
      Begin VB.Menu mnutax 
         Caption         =   "Ta&x Setting"
      End
      Begin VB.Menu mnuDepartment 
         Caption         =   "&Department Setting"
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu mnurestore 
         Caption         =   "&Restore all Settings"
      End
   End
   Begin VB.Menu mnuPaySlip 
      Caption         =   "&Payslip"
      Begin VB.Menu mnuPaySlipGenerate 
         Caption         =   "Generate Payslip"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuRPrintPayslip 
         Caption         =   "Re-Print Payslip"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuName 
         Caption         =   "All Payslips of one Employee"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuALl 
         Caption         =   "All Payslip Data"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCal 
         Caption         =   "&Calendar"
      End
      Begin VB.Menu mnuCalcu 
         Caption         =   "C&alculator"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
      Begin VB.Menu bar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDevelope 
         Caption         =   "&Developer"
      End
   End
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
Me.Show
Form1.Show vbModal
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Set con = Nothing
End Sub

Private Sub mnuadmin_Click()
MsgBox "Demo Version :: Project in IT - 11 ::", vbInformation
End Sub

Private Sub mnuALl_Click()
typeID = "c"
Load frmReprintPayslip
frmReprintPayslip.Show
End Sub

Private Sub mnuDsetup_Click()
Designation_form.Show
End Sub

Private Sub mnuCal_Click()
Form3.Show vbModal
End Sub

Private Sub mnuCalcu_Click()
On Error GoTo err
Shell "calc.exe", vbNormalFocus
Exit Sub
err:
    MsgBox "You don't have a Calculator installed in your computer.", vbExclamation, "System Error"
End Sub

Private Sub mnuDepartment_Click()
Department_form.Show vbModal
End Sub

Private Sub mnuDevelope_Click()
frmAbout.Show vbModal

End Sub

Private Sub mnuEmployee_Click()
Load frmEmpDetails
frmEmpDetails.Show

End Sub

Private Sub mnuExit_Click()

Dim reply01 As Integer
reply01 = MsgBox("This will terminate the application.Do you want to proceed?", vbExclamation + vbYesNo)
If reply01 = vbYes Then
End
Else
End If
End Sub

Private Sub mnuLeave_Click()
Leave_form.Show vbModal
End Sub

Private Sub mnuName_Click()
typeID = "b"
Load frmReprintPayslip
frmReprintPayslip.Show

End Sub

Private Sub mnuPaySlipGenerate_Click()
Load frmPaySlipGeneration
frmPaySlipGeneration.Show
End Sub

Private Sub mnuRPrintPayslip_Click()
typeID = "a"
Load frmReprintPayslip
frmReprintPayslip.Show
End Sub

Private Sub mnuSetDesig_Click()
Designation_form.Show vbModal
End Sub



Private Sub mnutax_Click()
tax_form.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 2: mnuEmployee_Click
    Case 3: mnutax_Click
    Case 4: mnuPaySlipGenerate_Click
    Case 6: mnuRPrintPayslip_Click
    Case 7: mnuName_Click
    Case 9: mnuALl_Click
    Case 10: mnuDevelope_Click
End Select

End Sub
