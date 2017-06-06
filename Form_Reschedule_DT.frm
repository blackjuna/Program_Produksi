VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Reschedule_DT 
   Caption         =   "Form Reschedule DT"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Filter Tanggal"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2655
      Begin MSComCtl2.DTPicker dtfilter 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   41739
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11668
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu pop_up_menu 
      Caption         =   "Pop Up Menu"
      Visible         =   0   'False
      Begin VB.Menu edit 
         Caption         =   "Edit"
      End
      Begin VB.Menu reschedule 
         Caption         =   "Reschedule"
      End
   End
   Begin VB.Menu refresh 
      Caption         =   "Refresh"
   End
End
Attribute VB_Name = "Form_Reschedule_DT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dtfilter_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    qry_filter = "select no_slip,no_so,jic,size,qty,finish_date,delivery_date,status,qty_pending,remarks_produksi from completion_slip where finish_date='" & Format(dtfilter.Value, "YYYY/mm/dd") & "'"
    Set rscompletion_slip = conn.Execute(qry_filter)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
    Set lst = ListView1.ListItems.Add
    lst.Text = rscompletion_slip!no_slip
    lst.SubItems(1) = rscompletion_slip!no_slip
    lst.SubItems(2) = rscompletion_slip!no_so
    lst.SubItems(3) = rscompletion_slip!jic
    lst.SubItems(4) = rscompletion_slip!Size
    lst.SubItems(5) = rscompletion_slip!qty
    lst.SubItems(6) = rscompletion_slip!finish_date
    lst.SubItems(7) = rscompletion_slip!delivery_date
    lst.SubItems(8) = rscompletion_slip!status
    lst.SubItems(9) = rscompletion_slip!qty_pending
    lst.SubItems(10) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
    rscompletion_slip.MoveNext
    Loop
    End With
    Call Warna_List
End If

End Sub

Private Sub Form_Load()
dtfilter.Value = Date + 3
Call db
Call SetLV
If rscompletion_slip.State = 1 Then rscompletion_slip.Close
'rscompletion_slip.Open "Select * from completion_slip ", conn
Call TplGrid
Call Warna_List
Set rscompletion_slip = Nothing
End Sub

Public Sub SetLV()
With ListView1
    .View = lvwReport
    .GridLines = True
    .MultiSelect = True
    .FullRowSelect = True
    .HotTracking = True
    .HoverSelection = True
    ' tambahkan kolom2 ke, , Judul,lebar,aligment
    .ColumnHeaders.Add 1, , "No Slip", 0
    .ColumnHeaders.Add 2, , "No Slip", 1700
    .ColumnHeaders.Add 3, , "No Sales Order", 1400
    .ColumnHeaders.Add 4, , "JIC", 2500
    .ColumnHeaders.Add 5, , "Size", 2200
    .ColumnHeaders.Add 6, , "Qty", 750
    .ColumnHeaders.Add 7, , "DT PPIC", 1000
    .ColumnHeaders.Add 8, , "Realisasi DT", 1100
    .ColumnHeaders.Add 9, , "Status", 1100
    .ColumnHeaders.Add 10, , "Qty Pending", 1100
    .ColumnHeaders.Add 11, , "Remarks Produksi", 3000
    .Width = 15000
End With
End Sub

Sub TplGrid()
    Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    lihat = "select no_slip,no_so,jic,size,qty,finish_date,delivery_date,status,qty_pending,remarks_produksi from completion_slip where finish_date='" & Format(dtfilter.Value, "YYYY/mm/dd") & "'"
    Set rscompletion_slip = conn.Execute(lihat)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
    Set lst = ListView1.ListItems.Add
    lst.Text = rscompletion_slip!no_slip
    lst.SubItems(1) = rscompletion_slip!no_slip
    lst.SubItems(2) = rscompletion_slip!no_so
    lst.SubItems(3) = rscompletion_slip!jic
    lst.SubItems(4) = rscompletion_slip!Size
    lst.SubItems(5) = rscompletion_slip!qty
    lst.SubItems(6) = rscompletion_slip!finish_date
    lst.SubItems(7) = rscompletion_slip!delivery_date
    lst.SubItems(8) = rscompletion_slip!status
    lst.SubItems(9) = rscompletion_slip!qty_pending
    lst.SubItems(10) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
    rscompletion_slip.MoveNext
    Loop
    End With
   
End Sub

Sub Warna_List()
Dim i As Long

For i = 1 To ListView1.ListItems.Count
If ListView1.ListItems(i).SubItems(8) = "Pending" Then 'Field Stok pada kolom 5
ListView1.ListItems(i).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(1).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(2).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(3).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(4).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(5).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(6).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(7).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(8).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(9).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems(10).ForeColor = vbRed
ElseIf ListView1.ListItems(i).SubItems(8) = "Partial" Then
ListView1.ListItems(i).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(1).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(2).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(3).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(4).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(5).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(6).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(7).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(8).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(9).ForeColor = vbGreen
ListView1.ListItems(i).ListSubItems(10).ForeColor = vbGreen
ElseIf ListView1.ListItems(i).SubItems(8) = "Closed" Then
ListView1.ListItems(i).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(1).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(2).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(3).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(4).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(5).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(6).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(7).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(8).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(9).ForeColor = vbBlue
ListView1.ListItems(i).ListSubItems(10).ForeColor = vbBlue

Else
ListView1.ListItems(i).ForeColor = vbBlack
ListView1.ListItems(i).ListSubItems(1).ForeColor = vbBlack
ListView1.ListItems(i).ListSubItems(2).ForeColor = vbBlack
ListView1.ListItems(i).ListSubItems(3).ForeColor = vbBlack
ListView1.ListItems(i).ListSubItems(4).ForeColor = vbBlack
ListView1.ListItems(i).ListSubItems(5).ForeColor = vbBlack
End If
Next

End Sub


Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then PopupMenu pop_up_menu
End Sub

Private Sub refresh_Click()
Call TplGrid
Call Warna_List

End Sub

Private Sub reschedule_Click()
ubah = "update completion_slip SET status='Pending' where no_slip='" & ListView1.SelectedItem.Text & "'"
Set rscompletion_slip = conn.Execute(ubah)
Call TplGrid
Call Form_Load
End Sub
