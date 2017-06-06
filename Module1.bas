Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection
Public rscode As New ADODB.Recordset
Public vread As New ADODB.Recordset
Public rscompletion_slip As New ADODB.Recordset
Public rsdata_mesin As New ADODB.Recordset
Public rsprint As New ADODB.Recordset

Public Sub db()
Set conn = New ADODB.Connection
Set rscode = New ADODB.Recordset
Set vread = New ADODB.Recordset
Set rscompletion_slip = New ADODB.Recordset
Set rsdata_mesin = New ADODB.Recordset
'koneksi = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=penjadwalan_produksi;Data Source =192.168.10.250, 1433;User Id=sa;Password=admin123"
koneksi = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=purchasing;Data Source =192.168.10.250, 1433;User Id=sa;Password=admin123"
conn.CursorLocation = adUseClient
conn.Open koneksi
End Sub

