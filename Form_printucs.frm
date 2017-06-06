VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_printucs 
   Caption         =   "Export To Excel"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Filter Export"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   2880
         TabIndex        =   9
         Top             =   120
         Width           =   5535
         Begin VB.CheckBox chk_dt_real 
            Caption         =   "Realisasi DT"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dt_real 
            Height          =   375
            Left            =   720
            TabIndex        =   11
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16449537
            CurrentDate     =   41739
         End
         Begin MSComCtl2.DTPicker dt_real2 
            Height          =   375
            Left            =   3360
            TabIndex        =   12
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16449537
            CurrentDate     =   41739
         End
         Begin VB.Label Label2 
            Caption         =   "To"
            Height          =   255
            Left            =   2880
            TabIndex        =   14
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "From"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.CommandButton btn_export 
         Caption         =   "EXPORT TO EXCEL"
         Height          =   495
         Left            =   3120
         TabIndex        =   7
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CheckBox chk_cb_status 
         Caption         =   "Status"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CheckBox chk_no_so 
         Caption         =   "No. Sales Order"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CheckBox chk_dt_ppic 
         Caption         =   "Delivery Time PPIC"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txt_no_so 
         Height          =   405
         Left            =   360
         TabIndex        =   2
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox cb_status 
         Height          =   315
         Left            =   3120
         TabIndex        =   1
         Top             =   1920
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dt_time_ppic 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   41739
      End
      Begin MSComctlLib.ProgressBar pgb 
         Height          =   300
         Left            =   0
         TabIndex        =   8
         Top             =   3240
         Visible         =   0   'False
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "Form_printucs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strstatus, strdtppic, strdtreal, strnoso, strquery, strchk1, strchk2, _
strchk3, strchk4, strchk, strsql As String
Dim lngctrlbrs As Long
Dim intValue As Integer
Dim ApExcel As Object


Private Sub btn_export_Click()
    ExportToExcel
End Sub

Private Sub chk_cb_status_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If chk_cb_status.Value = 1 Then cb_status.Enabled = True Else cb_status.Enabled = False
End Sub

Private Sub chk_dt_ppic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If chk_dt_ppic.Value = 1 Then dt_time_ppic.Enabled = True Else dt_time_ppic.Enabled = False
End Sub

Private Sub chk_dt_real_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If chk_dt_real.Value = 1 Then dt_real.Enabled = True: dt_real2.Enabled = True Else dt_real.Enabled = False: dt_real2.Enabled = False
End Sub

Private Sub chk_no_so_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If chk_no_so.Value = 1 Then txt_no_so.Enabled = True Else txt_no_so.Enabled = False
End Sub

Private Function ExportToExcel()
    Screen.MousePointer = vbHourglass
    
    If chk_dt_ppic.Value = 1 Then
        strchk1 = "I"
        strdtppic = "finish_date = '" & Format(dt_time_ppic.Value, "yyyy-mm-dd") & "'"
    Else
        strchk1 = "O"
        strdtppic = ""
    End If
    
    If chk_dt_real.Value = 1 Then
        strchk2 = "I"
        strdtreal = "delivery_date between '" & Format(dt_real.Value, "yyyy-mm-dd") & "' and '" & Format(dt_real2.Value, "yyyy-mm-dd") & "'"
    Else
        strchk2 = "O"
        strdtreal = ""
    End If
    
    If chk_no_so.Value = 1 Then
        strchk3 = "I"
        strnoso = "no_so like '%" & txt_no_so.Text & "%'"
    Else
        strchk3 = "O"
        strnoso = ""
    End If
    
    If chk_cb_status.Value = 1 Then
        strchk4 = "I"
        strstatus = "status like '%" & cb_status.Text & "%'"
    Else
        strchk4 = "O"
        strstatus = ""
    End If
    
    strchk = strchk1 & strchk2 & strchk3 & strchk4
    Select Case strchk
        Case "OOOO"
            strquery = "where deleted=0"
        Case "OOOI"
            strquery = "where " & strstatus & " and deleted=0"
        Case "OOII"
            strquery = "where " & strnoso & " and " & strstatus & " and deleted=0"
        Case "OIII"
            strquery = "where " & strdtreal & " and " & strnoso & " and " & strstatus & " and deleted=0"
        Case "IIII"
            strquery = "where " & strdtppic & " and " & strdtreal & " and " & strnoso & " and " & strstatus & " and deleted=0"
        Case "IIIO"
            strquery = "where " & strdtppic & " and " & strdtreal & " and " & strnoso & " and deleted=0"
        Case "IIOO"
            strquery = "where " & strdtppic & " and " & strdtreal & " and deleted=0"
        Case "IOOO"
            strquery = "where " & strdtppic & " and deleted=0"
        Case "OOIO"
            strquery = "where " & strnoso & " and deleted=0"
        Case "OIIO"
            strquery = "where " & strdtreal & " and " & strnoso & " and deleted=0"
        Case "OIOO"
            strquery = "where " & strdtreal & " and deleted=0"
        Case "OIOI"
            strquery = "where " & strdtreal & " and " & strstatus & " and deleted=0"
        Case "IOOI"
            strquery = "where " & strdtppic & " and " & strstatus & " and deleted=0"
        Case "IOIO"
            strquery = "where " & strdtppic & " and " & strnoso & " and deleted=0"
        Case "IOII"
            strquery = "where " & strdtppic & " and " & strnoso & " and " & strstatus & " and deleted=0"
        Case "IIOI"
            strquery = "where " & strdtppic & " and " & strdtreal & " and " & strstatus & " and deleted=0"
    End Select
    
    Set rsprint = New ADODB.Recordset
    
    If rsprint.State = 1 Then rsprint.Close
    strsql = "Select no_slip,no_so,no_part,jic,size,qty,finish_date,delivery_date," & _
        "status,qty_pending,remarks_produksi from completion_slip " & _
        "" & strquery & " "
    rsprint.Open strsql, conn, adOpenDynamic, adLockOptimistic
    
    If rsprint.EOF Then
        MsgBox "Data Tidak Ada."
        rsprint.Close
        Set rsprint = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    intValue = 0
    pgb.Min = 0
    pgb.Max = IIf(rsprint.RecordCount < 1, 2, rsprint.RecordCount)
    pgb.Visible = True

    Set ApExcel = CreateObject("Excel.application")
    ApExcel.Visible = False
    ApExcel.Workbooks.Add
    
    setTitle 1, 1, "LIST COMPLETION SLIP", 20
    setColTitle 3, 1, "NO SLIP"
    setColTitle 3, 2, "NO SALES ORDER"
    setColTitle 3, 3, "NO PART"
    setColTitle 3, 4, "JIC"
    setColTitle 3, 5, "SIZE"
    setColTitle 3, 6, "QTY"
    setColTitle 3, 7, "DT PPIC"
    setColTitle 3, 8, "REALISASI DT"
    setColTitle 3, 9, "STATUS"
    setColTitle 3, 10, "QTY PENDING"
    setColTitle 3, 11, "REMARKS PRODUKSI"

    lngctrlbrs = 4
    rsprint.MoveFirst
    Do While Not rsprint.EOF
        intValue = intValue + 1
        pgb.Value = intValue
        
        
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 1).Value = "'" & rsprint!no_slip
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 1).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 2).Value = "'" & rsprint!no_so
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 2).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 3).Value = "'" & rsprint!no_part
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 3).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 4).Value = "'" & rsprint!jic
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 4).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 5).Value = "'" & rsprint!Size
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 5).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 6).Value = rsprint!qty
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 6).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 7).Value = "'" & rsprint!finish_date
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 7).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 8).Value = "'" & rsprint!delivery_date
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 8).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 9).Value = "'" & rsprint!Status
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 9).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 10).Value = rsprint!qty_pending
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 10).Font.Size = 10
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 11).Value = "'" & rsprint!remarks_produksi
        ApExcel.ActiveSheet.Cells(lngctrlbrs, 11).Font.Size = 10
        rsprint.MoveNext
        lngctrlbrs = lngctrlbrs + 1
    Loop
    
    lngctrlbrs = lngctrlbrs - 1
    ApExcel.Range(ApExcel.ActiveSheet.Cells(3, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 11)).Borders(1).LineStyle = 1
    ApExcel.Range(ApExcel.ActiveSheet.Cells(3, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 11)).Borders(2).LineStyle = 1
    ApExcel.Range(ApExcel.ActiveSheet.Cells(3, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 11)).Borders(3).LineStyle = 1
    ApExcel.Range(ApExcel.ActiveSheet.Cells(3, 1), ApExcel.ActiveSheet.Cells(lngctrlbrs, 11)).Borders(4).LineStyle = 1

    
    ApExcel.Columns.AutoFit
    ApExcel.Columns(1).ColumnWidth = 20
    ApExcel.Visible = True
    
    pgb.Visible = False
    Screen.MousePointer = vbDefault
    
    Set ApExcel = Nothing
    
    rsprint.Close
    Set rsprint = Nothing
    
    
End Function


Private Function setColTitle(lngbaris As Long, lngkolom As Long, strValue As String)
    ApExcel.Cells(lngbaris, lngkolom).Font.Size = 10
    ApExcel.Cells(lngbaris, lngkolom).Font.Bold = True
    ApExcel.Cells(lngbaris, lngkolom).Value = strValue
    ApExcel.Cells(lngbaris, lngkolom).Interior.ColorIndex = 36
    ApExcel.Cells(lngbaris, lngkolom).WrapText = False
End Function

Private Function setTitle(lngbaris As Long, lngkolom As Long, strValue As String, Optional intFontSize As Integer = 15)
    ApExcel.Cells(lngbaris, lngkolom).Font.Size = intFontSize
    ApExcel.Cells(lngbaris, lngkolom).Font.Bold = True
    ApExcel.Cells(lngbaris, lngkolom).Value = strValue
    ApExcel.Cells(lngbaris, lngkolom).WrapText = False
End Function

Private Function AddStatus()
    Set rscode = New ADODB.Recordset
    
    strsql = "Select Distinct status from completion_slip where deleted=0"
    rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
    
    If Not rscode.EOF Then
        rscode.MoveFirst
        Do While Not rscode.EOF
            cb_status.AddItem rscode!Status
            rscode.MoveNext
        Loop
        rscode.Close
        Set rscode = Nothing
    Else
        rscode.Close
        Set rscode = Nothing
    End If
End Function

Private Sub dt_real_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dt_real2.SetFocus
End Sub

Private Sub Form_Load()
    Form_printucs.Top = (Screen.Height - Form_printucs.Height) \ 2
    Form_printucs.Left = (Screen.Width - Form_printucs.Width) \ 2
    
    dt_time_ppic.Enabled = False
    dt_real.Enabled = False
    dt_real.Value = Date
    dt_real2.Value = Date
    dt_real2.Enabled = False
    txt_no_so.Enabled = False
    cb_status.Enabled = False
    AddStatus
End Sub
