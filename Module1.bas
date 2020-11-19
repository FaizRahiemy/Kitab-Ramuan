Attribute VB_Name = "Module1"
Option Explicit
Public Sql As String
Public db As New ADODB.Connection
Public rs As New ADODB.Recordset
Public Const AbuAbu = &H8000000F
Public Const Putih = &H80000005
Public Sub Koneksi()
Set db = New ADODB.Connection
db.Provider = "Microsoft.Jet.OLEDB.4.0 "
db.Open (App.Path & "\Database.mdb")
Set rs = New ADODB.Recordset
End Sub
