VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Cert 
   Caption         =   "Form Add Certificate"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmd_deselect 
      Caption         =   "Unselect All"
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmd_select 
      Caption         =   "Select All"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Completion Slip"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox tstatus 
         Height          =   285
         Left            =   5400
         TabIndex        =   20
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox tqty 
         Height          =   285
         Left            =   5400
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox tsize 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox tjic 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox tno_so 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox tno_slip 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox tremarks 
         Height          =   525
         Left            =   5400
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtdeliverydate 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   41739
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         Height          =   195
         Left            =   3840
         TabIndex        =   14
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "JIC"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No SO"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   3840
         TabIndex        =   10
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Slip"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Remarks Produksi"
         Height          =   195
         Left            =   3840
         TabIndex        =   8
         Top             =   1080
         Width           =   1290
      End
   End
   Begin MSComctlLib.ListView lv_add_cert 
      Height          =   2775
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form_Cert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strid As String
Public strsql As String

Private Sub Command3_Click()

End Sub

Private Sub cmd_deselect_Click()
    For i = 1 To lv_add_cert.ListItems.Count
        lv_add_cert.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmd_save_Click()
    If vread.State = 1 Then vread.Close
    strsql = "select id from completion_slip where no_slip='" & strid & "' "
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not vread.EOF Then
        strsql = "delete from cs_files where id_cs='" & vread!id & "'"
        conn.Execute (strsql)
        For i = 1 To lv_add_cert.ListItems.Count
            If lv_add_cert.ListItems(i).Checked = True Then
                strsql = "insert into cs_files (id_cs,id_certificate_files,deleted) values " & _
                    "('" & vread!id & "','" & lv_add_cert.ListItems(i) & "',0)"
                conn.Execute (strsql)
            End If
        Next
        MsgBox "Certificate Sudah tersimpan", vbOKOnly + vbInformation, "Informasi"
    End If
End Sub

Private Sub cmd_select_Click()
    For i = 1 To lv_add_cert.ListItems.Count
        lv_add_cert.ListItems(i).Checked = True
    Next
End Sub

Private Sub Form_Load()
    'Call SetLV
    Call EnabledAll(False)
    Call SetLVCert
    strid = Form_Update_CS.strid
    
    If vread.State = 1 Then vread.Close
    strsql = "select * from completion_slip where no_slip='" & strid & "' "
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not vread.EOF Then
        tno_slip.Text = IIf(IsNull(vread!no_slip), "", vread!no_slip)
        tno_so.Text = IIf(IsNull(vread!no_so), "", vread!no_so)
        tjic.Text = IIf(IsNull(vread!jic), "", vread!jic)
        tsize.Text = IIf(IsNull(vread!Size), "", vread!Size)
        tqty.Text = IIf(IsNull(vread!qty), "", vread!Size)
        dtdeliverydate.Value = IIf(IsNull(vread!delivery_date), "", vread!delivery_date)
        tstatus.Text = IIf(IsNull(vread!status), "", vread!status)
        tremarks.Text = IIf(IsNull(vread!remarks_produksi), "", vread!remarks_produksi)
    End If
    
    strsql = "Select * from certificate_files where deleted=0"
    Call LoadListView(strsql)
End Sub

Public Sub LoadListView(strsql As String)
    Dim lst As ListItem
    lv_add_cert.ListItems.Clear
    If vread.State = 1 Then vread.Close
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    
    If Not vread.EOF Then
        vread.MoveFirst
        i = 1
        Do While Not vread.EOF
            Set lst = lv_add_cert.ListItems.Add
            
            If rscode.State = 1 Then rscode.Close
            strsql = "select no_slip,no_so,no_part,jic,size,marking_stamp_or,marking_stamp_ir,qty,completion_slip.note as cs_note,customer,status," & _
                "certificate_files.file_name,certificate_files.note as cert_note,certificate_files.title,certificate_files.id as id_cert from completion_slip " & _
                "LEFT JOIN cs_files ON cs_files.id_cs  =completion_slip.id " & _
                "LEFT JOIN certificate_files on cs_files.id_certificate_files=certificate_files.id " & _
                "where completion_slip.no_slip='" & strid & "' and " & _
                "certificate_files.id='" & Format(IIf(IsNull(vread!id), "", vread!id)) & "'"
            rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
            
            lst.Text = Format(IIf(IsNull(vread!id), "", vread!id))
            
            If Not rscode.EOF Then lv_add_cert.ListItems.Item(i).Checked = True
            lst.SubItems(1) = Format(IIf(IsNull(vread!Date), "", Format(vread!Date, "dd-mm-yyyy")))
            lst.SubItems(2) = Format(IIf(IsNull(vread!file_name), "", vread!file_name))
            lst.SubItems(3) = Format(IIf(IsNull(vread!Title), "", vread!Title))
            lst.SubItems(4) = Format(IIf(IsNull(vread!note), "", vread!note))
            i = i + 1
            vread.MoveNext
        Loop
        vread.Close
    End If
End Sub

Public Sub SetLVCert()
    With lv_add_cert
        .View = lvwReport
        .GridLines = True
        .MultiSelect = True
        .FullRowSelect = True
        .HotTracking = True
        .HoverSelection = True
        ' tambahkan kolom2 ke, , Judul,lebar,aligment
        .ColumnHeaders.Add 1, , "", 300
        .ColumnHeaders.Add 2, , "Incoming Date", 1300
        .ColumnHeaders.Add 3, , "File Name", 1800
        .ColumnHeaders.Add 4, , "title", 2500
        .ColumnHeaders.Add 5, , "Note", 4500
    End With
End Sub

Public Sub EnabledAll(status As Boolean)
    For Each a In Me
        If TypeOf a Is TextBox Then a.Enabled = status
        If TypeOf a Is DTPicker Then a.Enabled = status
    Next a
End Sub

Private Sub lv_add_cert_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index - 1 <> CURR_COL Then
        lv_add_cert.SortOrder = 0
    Else
        lv_add_cert.SortOrder = Abs(lv_add_cert.SortOrder - 1)
    End If
    
    lv_add_cert.SortKey = ColumnHeader.Index - 1
    lv_add_cert.Sorted = True
    CURR_COL = ColumnHeader.Index - 1
End Sub
