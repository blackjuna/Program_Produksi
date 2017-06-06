VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Update_CS 
   Caption         =   "Form Update CS"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Certificate"
      Height          =   6015
      Left            =   11400
      TabIndex        =   41
      Top             =   3480
      Width           =   3615
      Begin MSComctlLib.ListView lv_cert 
         Height          =   5535
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   9763
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5895
      Left            =   120
      TabIndex        =   32
      Top             =   3600
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10398
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filter Order"
      Height          =   3135
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   6615
      Begin VB.TextBox txt_no_part 
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Realisasi DT"
         Height          =   1335
         Left            =   2280
         TabIndex        =   33
         Top             =   360
         Width           =   4215
         Begin MSComCtl2.DTPicker dtreal 
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16515073
            CurrentDate     =   41739
         End
         Begin MSComCtl2.DTPicker dtreal2 
            Height          =   375
            Left            =   2160
            TabIndex        =   35
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16515073
            CurrentDate     =   41739
         End
         Begin VB.Label Label18 
            Caption         =   "To"
            Height          =   255
            Left            =   2160
            TabIndex        =   38
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "From"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.ComboBox filter_status 
         Height          =   315
         Left            =   2400
         TabIndex        =   31
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox tsales_order 
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtfilter 
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   41739
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "No. Part"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   2400
         TabIndex        =   30
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "No. Sales Order"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Time PPIC"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Completion Slip"
      Height          =   3135
      Left            =   6840
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.CommandButton cmd_certificate 
         Caption         =   "Add Cert"
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox tremarks 
         Height          =   525
         Left            =   5400
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox tno_slip 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox tqtypending 
         Height          =   285
         Left            =   5400
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cstatus 
         Height          =   315
         Left            =   5400
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox tno_so 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox tjic 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox tsize 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox tqty 
         Height          =   285
         Left            =   5400
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox tqty_aktual 
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox theat_number 
         Height          =   285
         Left            =   5400
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox tcert_number 
         Height          =   285
         Left            =   5400
         TabIndex        =   1
         Top             =   2040
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtdeliverydate 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   41739
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Remarks Produksi"
         Height          =   195
         Left            =   3840
         TabIndex        =   24
         Top             =   2400
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Slip"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   3840
         TabIndex        =   21
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Qty Pending"
         Height          =   195
         Left            =   3840
         TabIndex        =   20
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No SO"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "JIC"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         Height          =   195
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Date"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Qty Actual"
         Height          =   195
         Left            =   3840
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Heat Number"
         Height          =   195
         Left            =   3840
         TabIndex        =   13
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Cert Number"
         Height          =   195
         Left            =   3840
         TabIndex        =   12
         Top             =   2040
         Width           =   885
      End
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
   Begin VB.Menu reschedule_dt 
      Caption         =   "Reschedule DT"
   End
   Begin VB.Menu exp_excel 
      Caption         =   "Export To Excel"
   End
End
Attribute VB_Name = "Form_Update_CS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strid As String
Public strsql As String

Private Function nomor_part()
    Dim strsql As String
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    strsql = "Select no_part from completion_slip"
    rscompletion_slip.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If rscompletion_slip.EOF Then
        rscompletion_slip.Close
    Else
        rscompletion_slip.MoveFirst
        Do While Not rscompletion_slip.EOF
            cb_no_part.AddItem rscompletion_slip!no_part
            rscompletion_slip.MoveNext
        Loop
    End If
End Function

Sub Warna_List()
Dim i As Long

For i = 1 To ListView1.ListItems.Count
If ListView1.ListItems(i).SubItems(9) = "Pending" Then 'Field Stok pada kolom 5
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
ListView1.ListItems(i).ListSubItems(11).ForeColor = vbRed
ElseIf ListView1.ListItems(i).SubItems(9) = "Partial" Then
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
ListView1.ListItems(i).ListSubItems(11).ForeColor = vbGreen
ElseIf ListView1.ListItems(i).SubItems(9) = "Closed" Then
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
ListView1.ListItems(i).ListSubItems(11).ForeColor = vbBlue
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
    .ColumnHeaders.Add 4, , "No Part", 1700
    .ColumnHeaders.Add 5, , "JIC", 2500
    .ColumnHeaders.Add 6, , "Size", 2200
    .ColumnHeaders.Add 7, , "Qty", 750
    .ColumnHeaders.Add 8, , "DT PPIC", 1000
    .ColumnHeaders.Add 9, , "Realisasi DT", 1100
    .ColumnHeaders.Add 10, , "Status", 1100
    .ColumnHeaders.Add 11, , "Qty Pending", 1100
    .ColumnHeaders.Add 12, , "Remarks Produksi", 3000
    .ColumnHeaders.Add 13, , "id", 0
    '.Width = 15000
End With
End Sub
Sub TplGrid()
    Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    lihat = "select no_slip,no_so,no_part,jic,size,qty,finish_date,delivery_date, " & _
        "status,qty_pending, remarks_produksi,id from completion_slip"
    Set rscompletion_slip = conn.Execute(lihat)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
        Set lst = ListView1.ListItems.Add
        lst.Text = rscompletion_slip!no_slip
        lst.SubItems(1) = rscompletion_slip!no_slip
        lst.SubItems(2) = rscompletion_slip!no_so
        lst.SubItems(3) = rscompletion_slip!no_part
        lst.SubItems(4) = rscompletion_slip!jic
        lst.SubItems(5) = rscompletion_slip!Size
        lst.SubItems(6) = rscompletion_slip!qty
        lst.SubItems(7) = rscompletion_slip!finish_date
        lst.SubItems(8) = rscompletion_slip!delivery_date
        lst.SubItems(9) = rscompletion_slip!status
        lst.SubItems(10) = IIf(IsNull(rscompletion_slip!qty_pending), 0, rscompletion_slip!qty_pending)
        lst.SubItems(11) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
        lst.SubItems(12) = rscompletion_slip!id
    rscompletion_slip.MoveNext
    Loop
    End With
   
End Sub

Sub update()
ubah = "UPDATE completion_slip SET delivery_date='" & Format(dtdeliverydate.Value, "YYYY/mm/dd") & "', status='" & cstatus.Text & "', qty_pending='" & tqtypending.Text & "', heat_number='" & theat_number.Text & "', cert_number='" & tcert_number.Text & "',remarks_produksi='" & tremarks.Text & "' where no_slip='" & tno_slip.Text & "'"
    Set rscompletion_slip = conn.Execute(ubah)
    'tampilgrid
    Call bersih
    Call Form_Activate
End Sub

Private Sub crefresh_Click()

End Sub

Private Sub cmd_certificate_Click()
    Form_Cert.Show vbModal
End Sub

Private Sub cstatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If cstatus.Text = "Closed" Then
        tqty_aktual.Text = 0
        tqtypending.Text = 0
        theat_number.SetFocus
    Else
        tqty_aktual.SetFocus
    End If
End If
End Sub

Private Sub dtdeliverydate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cstatus.SetFocus
End If
End Sub

Private Sub dtfilter_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    qry_filter = "select id,no_slip,no_part,no_so,jic,size,qty,finish_date,delivery_date,status,qty_pending,remarks_produksi from completion_slip where finish_date='" & Format(dtfilter.Value, "YYYY-mm-dd") & "'"
    Set rscompletion_slip = conn.Execute(qry_filter)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
        Set lst = ListView1.ListItems.Add
        lst.Text = rscompletion_slip!no_slip
        lst.SubItems(1) = rscompletion_slip!no_slip
        lst.SubItems(2) = rscompletion_slip!no_so
        lst.SubItems(3) = rscompletion_slip!no_part
        lst.SubItems(4) = rscompletion_slip!jic
        lst.SubItems(5) = rscompletion_slip!Size
        lst.SubItems(6) = rscompletion_slip!qty
        lst.SubItems(7) = rscompletion_slip!finish_date
        lst.SubItems(8) = rscompletion_slip!delivery_date
        lst.SubItems(9) = rscompletion_slip!status
        lst.SubItems(10) = IIf(IsNull(rscompletion_slip!qty_pending), "", rscompletion_slip!qty_pending)
        lst.SubItems(11) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
        lst.SubItems(12) = rscompletion_slip!id
        rscompletion_slip.MoveNext
        Loop
    End With
    Call Warna_List
End If
End Sub

Private Sub dtreal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtreal2.SetFocus
End Sub

Private Sub dtreal2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    qry_filter = "select id,no_slip,no_part,no_so,jic,size,qty,finish_date," & _
        "delivery_date,status,qty_pending,remarks_produksi from completion_slip " & _
        "where delivery_date between '" & Format(dtreal.Value, "yyyy-mm-dd") & "' " & _
        "and '" & Format(dtreal2.Value, "yyyy-mm-dd") & "' "
    Set rscompletion_slip = conn.Execute(qry_filter)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
    Set lst = ListView1.ListItems.Add
    lst.SubItems(1) = rscompletion_slip!no_slip
    lst.SubItems(2) = rscompletion_slip!no_so
    lst.SubItems(3) = rscompletion_slip!no_part
    lst.SubItems(4) = rscompletion_slip!jic
    lst.SubItems(5) = rscompletion_slip!Size
    lst.SubItems(6) = rscompletion_slip!qty
    lst.SubItems(7) = rscompletion_slip!finish_date
    lst.SubItems(8) = rscompletion_slip!delivery_date
    lst.SubItems(9) = rscompletion_slip!status
    lst.SubItems(10) = rscompletion_slip!qty_pending
    lst.SubItems(11) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
    lst.SubItems(12) = rscompletion_slip!id
    rscompletion_slip.MoveNext
    Loop
    End With
    Call Warna_List
End If
End Sub

Private Sub edit_Click()
'Call db
cari = "select no_slip,no_so,jic,size,qty,delivery_date,status,qty_pending,remarks_produksi from completion_slip where no_slip='" & ListView1.SelectedItem.Text & "'"
Set rscompletion_slip = conn.Execute(cari)
    If Not rscompletion_slip.EOF Then
        tno_slip.Text = rscompletion_slip.Fields("no_slip")
        tno_so.Text = rscompletion_slip.Fields("no_so")
        tjic.Text = rscompletion_slip.Fields("jic")
        tsize.Text = rscompletion_slip.Fields("size")
        tqty.Text = rscompletion_slip.Fields("qty")
        dtdeliverydate.Value = rscompletion_slip.Fields("delivery_date")
        cstatus.Text = rscompletion_slip.Fields("status")
        tqtypending.Text = IIf(IsNull(rscompletion_slip.Fields("qty_pending")), "", rscompletion_slip.Fields("qty_pending"))
        tremarks.Text = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
    End If
tno_so.Enabled = False
tjic.Enabled = False
tsize.Enabled = False
End Sub

Private Sub exp_excel_Click()
    Form_printucs.Show vbModal
End Sub

Private Sub filter_status_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    qry_status = "select id,no_slip,no_part,no_so,jic,size,qty,finish_date,delivery_date,status,qty_pending,remarks_produksi from completion_slip where status='" & filter_status.Text & "'"
    Set rscompletion_slip = conn.Execute(qry_status)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
    Set lst = ListView1.ListItems.Add
    lst.Text = rscompletion_slip!no_slip
    lst.SubItems(1) = rscompletion_slip!no_slip
    lst.SubItems(2) = rscompletion_slip!no_so
    lst.SubItems(3) = rscompletion_slip!no_part
    lst.SubItems(4) = rscompletion_slip!jic
    lst.SubItems(5) = rscompletion_slip!Size
    lst.SubItems(6) = rscompletion_slip!qty
    lst.SubItems(7) = rscompletion_slip!finish_date
    lst.SubItems(8) = rscompletion_slip!delivery_date
    lst.SubItems(9) = rscompletion_slip!status
    lst.SubItems(10) = rscompletion_slip!qty_pending
    lst.SubItems(11) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
    lst.SubItems(12) = rscompletion_slip!id
    rscompletion_slip.MoveNext
    Loop
    End With
    Call Warna_List
End If

End Sub

Private Sub Form_Activate()
cstatus.AddItem "Partial"
cstatus.AddItem "Pending"
cstatus.AddItem "Closed"
filter_status.AddItem "On Going"
filter_status.AddItem "Partial"
filter_status.AddItem "Pending"
filter_status.AddItem "Closed"
dtdeliverydate.Value = Date
dtreal.Value = Date
dtreal2.Value = Date
dtfilter.Value = Date
tno_slip.SetFocus
End Sub

Private Sub Form_Load()
Call db
Call SetLV
Call SetLVCert
If rscompletion_slip.State = 1 Then rscompletion_slip.Close
Call TplGrid
Call Warna_List
'Call nomor_part
Set rscompletion_slip = Nothing
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    strid = ListView1.SelectedItem.Text
    strsql = "select id,id_cs,id_certificate_files from cs_files where id_cs='" & ListView1.SelectedItem.SubItems(12) & "' "
    Call LoadListView(strsql)
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then PopupMenu pop_up_menu
End Sub

Private Sub lv_cert_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index - 1 <> CURR_COL Then
        lv_cert.SortOrder = 0
    Else
        lv_cert.SortOrder = Abs(lv_cert.SortOrder - 1)
    End If
    
    lv_cert.SortKey = ColumnHeader.Index - 1
    lv_cert.Sorted = True
    CURR_COL = ColumnHeader.Index - 1
End Sub

Private Sub refresh_Click()
Call TplGrid
Call Warna_List
Call bersih

For Each a In Me
If TypeOf a Is TextBox Then a.Text = ""
Next a

tno_slip.SetFocus
End Sub

Private Sub reschedule_Click()
ubah = "update completion_slip SET status='Pending' where no_slip='" & ListView1.SelectedItem.Text & "'"
Set rscompletion_slip = conn.Execute(ubah)
Call TplGrid
Call Form_Load
End Sub

Private Sub reschedule_dt_Click()
Form_Reschedule_DT.Show
End Sub

Private Sub tcert_number_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tremarks.SetFocus
End If

End Sub

Private Sub theat_number_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tcert_number.SetFocus
End If
End Sub

Private Sub tno_slip_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    'Call db
    cari = "select no_so,jic,size,qty,status from completion_slip where no_slip='" & tno_slip.Text & "'"
    Set rscompletion_slip = conn.Execute(cari)
        If Not rscompletion_slip.EOF Then
            tno_so.Text = rscompletion_slip.Fields("no_so")
            tjic.Text = rscompletion_slip.Fields("jic")
            tsize.Text = rscompletion_slip.Fields("size")
            tqty.Text = rscompletion_slip.Fields("qty")
            cstatus.Text = rscompletion_slip.Fields("status")
            dtdeliverydate.SetFocus
        End If
    tno_so.Enabled = False
    tjic.Enabled = False
    tsize.Enabled = False
End If
    
End Sub

Private Sub tqty_aktual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tqtypending.Text = Val(tqty.Text) - Val(tqty_aktual.Text)
    theat_number.SetFocus
End If

End Sub

Private Sub tqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cstatus.SetFocus
End If
End Sub

Sub bersih()
For Each a In Me
    If TypeOf a Is TextBox Then a.Text = ""
    If TypeOf a Is ComboBox Then a.Text = ""
Next a

End Sub

Private Sub tremarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call update
    tno_slip.SetFocus
    Call bersih
End If

End Sub

Private Sub tsales_order_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    qry_sales = "select id,no_slip,no_part,no_so,jic,size,qty,finish_date," & _
        "delivery_date,status,qty_pending,remarks_produksi from completion_slip where no_so like '%" & tsales_order.Text & "%'"
    Set rscompletion_slip = conn.Execute(qry_sales)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
    Set lst = ListView1.ListItems.Add
    lst.Text = rscompletion_slip!no_slip
    lst.SubItems(1) = rscompletion_slip!no_slip
    lst.SubItems(2) = rscompletion_slip!no_so
    lst.SubItems(3) = rscompletion_slip!no_part
    lst.SubItems(4) = rscompletion_slip!jic
    lst.SubItems(5) = rscompletion_slip!Size
    lst.SubItems(6) = rscompletion_slip!qty
    lst.SubItems(7) = rscompletion_slip!finish_date
    lst.SubItems(8) = rscompletion_slip!delivery_date
    lst.SubItems(9) = rscompletion_slip!status
    lst.SubItems(10) = IIf(IsNull(rscompletion_slip!qty_pending), "", rscompletion_slip!qty_pending)
    lst.SubItems(11) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
    lst.SubItems(12) = rscompletion_slip!id
    rscompletion_slip.MoveNext
    Loop
    End With
    Call Warna_List
End If
End Sub

Private Sub txt_no_part_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    qry_sales = "select id,no_slip,no_part,no_so,jic,size,qty,finish_date," & _
        "delivery_date,status,qty_pending,remarks_produksi from completion_slip " & _
        "where no_part like '%" & txt_no_part.Text & "%'"
    Set rscompletion_slip = conn.Execute(qry_sales)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
    Set lst = ListView1.ListItems.Add
    lst.Text = rscompletion_slip!no_slip
    lst.SubItems(1) = rscompletion_slip!no_slip
    lst.SubItems(2) = rscompletion_slip!no_so
    lst.SubItems(3) = rscompletion_slip!no_part
    lst.SubItems(4) = rscompletion_slip!jic
    lst.SubItems(5) = rscompletion_slip!Size
    lst.SubItems(6) = rscompletion_slip!qty
    lst.SubItems(7) = rscompletion_slip!finish_date
    lst.SubItems(8) = rscompletion_slip!delivery_date
    lst.SubItems(9) = rscompletion_slip!status
    lst.SubItems(10) = rscompletion_slip!qty_pending
    lst.SubItems(11) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
    lst.SubItems(12) = rscompletion_slip!id
    rscompletion_slip.MoveNext
    Loop
    End With
    Call Warna_List
End If
End Sub

Public Sub SetLVCert()
    With lv_cert
        .View = lvwReport
        .GridLines = True
        .MultiSelect = True
        .FullRowSelect = True
        .HotTracking = True
        .HoverSelection = True
        ' tambahkan kolom2 ke, , Judul,lebar,aligment
        .ColumnHeaders.Add 1, , "", 0
        .ColumnHeaders.Add 2, , "Incoming Date", 1300
        .ColumnHeaders.Add 3, , "File Name", 1800
        .ColumnHeaders.Add 4, , "title", 2500
        .ColumnHeaders.Add 5, , "Note", 4500
    End With
End Sub

Public Sub LoadListView(strsql As String)
    Dim lst As ListItem
    lv_cert.ListItems.Clear
    
    If vread.State = 1 Then vread.Close
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not vread.EOF Then
        vread.MoveFirst
        Do While Not vread.EOF
            If rscode.State = 1 Then rscode.Close
            strsql = "Select id,date,file_name,title,note from certificate_files where id='" & vread!id_certificate_files & "'"
            rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
            If Not rscode.EOF Then
                Do While Not rscode.EOF
                    Set lst = lv_cert.ListItems.Add
                    lst.Text = Format(IIf(IsNull(rscode!id), "", rscode!id))
                    lst.SubItems(1) = Format(IIf(IsNull(rscode!Date), "", Format(rscode!Date, "dd-mm-yyyy")))
                    lst.SubItems(2) = Format(IIf(IsNull(rscode!file_name), "", rscode!file_name))
                    lst.SubItems(3) = Format(IIf(IsNull(rscode!Title), "", rscode!Title))
                    lst.SubItems(4) = Format(IIf(IsNull(rscode!note), "", rscode!note))
                    rscode.MoveNext
                Loop
            End If
            vread.MoveNext
        Loop
        vread.Close
    End If
End Sub

