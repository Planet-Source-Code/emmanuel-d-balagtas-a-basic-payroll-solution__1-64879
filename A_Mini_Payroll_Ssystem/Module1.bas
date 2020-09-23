Attribute VB_Name = "Module1"
Option Explicit
Public com As Command
Public con As Connection
Public rs As Recordset
Public rs1 As Recordset
Public rs11 As Recordset
Public rs2 As Recordset
Public rs3 As Recordset
Public rs4 As Recordset
Public typeID As String
Public strPath As String



Public Sub prcConnection()
Set com = New Command
Set con = New Connection
Set rs = New Recordset
Set rs11 = New Recordset
Set rs1 = New Recordset
Set rs2 = New Recordset
Set rs3 = New Recordset
Set rs4 = New Recordset

con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Payroll.mdb"
con.Open
com.ActiveConnection = con

End Sub

Public Function prcCallPaySlip(pays As String)
On Error GoTo err
DataEnvironment1.Command1 (Trim(pays))
rptPaySlip.Show
Exit Function
err:
Unload DataEnvironment1
DataEnvironment1.Command1 (Trim(pays))
rptPaySlip.Show
End Function
Public Function prcCallempPaySlip(pays As String)
On Error GoTo err
DataEnvironment1.Command2 (Trim(pays))
rptEmpSlips.Show
Exit Function
err:
Unload DataEnvironment1
DataEnvironment1.Command2 (Trim(pays))
rptEmpSlips.Show
End Function

Public Sub prcCallAllEmp(ToDate As String, fromDate As String)
On Error GoTo err
Call DataEnvironment1.comALL(ToDate, fromDate)
Load frmAllEmp
frmAllEmp.Show
Exit Sub
err:
Unload DataEnvironment1
'DataEnvironment1.Command2 (Trim(pays))
rptEmpSlips.Show
End Sub


