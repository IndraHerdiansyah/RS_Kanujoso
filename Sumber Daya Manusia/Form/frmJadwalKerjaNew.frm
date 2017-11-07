VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJadwalKerjaNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Jadwal Kerja Pegawai"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmJadwalKerjaNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   9255
   Begin MSComctlLib.ListView lvJadwalKerja 
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   68878337
      CurrentDate     =   40100
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   375
      Left            =   3120
      Picture         =   "frmJadwalKerjaNew.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Refresh Jadwal"
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid fgJadwalKerja 
      Height          =   3735
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6588
      _Version        =   393216
      FocusRect       =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkTabelJadwal 
      Caption         =   "&Munculkan Tabel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8040
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8505
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   16272
            Text            =   "F1 - Cetak"
            TextSave        =   "F1 - Cetak"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOtomatis 
      Caption         =   "&Jadwal Automatis"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox txtnamapegawai 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5520
      TabIndex        =   10
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox txtidpegawai 
      Height          =   285
      Left            =   7680
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtharilibur 
      Height          =   285
      Left            =   5760
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtnamatgl 
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtkdtgl 
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid dgJadwalKerja 
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcTempatBertugas 
      Height          =   330
      Left            =   5520
      TabIndex        =   1
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcShiftKerja 
      Height          =   315
      Left            =   5520
      TabIndex        =   4
      Top             =   2520
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cari Nama Pegawai :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   11
      Top             =   3120
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Shift Kerja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   5
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ruangan Tempat Bertugas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   2
      Top             =   1440
      Width           =   1920
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7440
      Picture         =   "frmJadwalKerjaNew.frx":0E14
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmJadwalKerjaNew.frx":2302
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmJadwalKerjaNew.frx":4CC3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmJadwalKerjaNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intRowLama As Integer, intColLama As Integer
Private intRowCur As Integer, intColCur As Integer
Private blnHariLibur As Boolean
Private intTglLibur As Integer

Private Sub chkTabelJadwal_Click()
    If Me.chkTabelJadwal.Value = 1 Then
        Call subSetFGJadwalKerja
        Call subIsiFgJadwalKerja
        Me.dcShiftKerja.Enabled = False
        Me.fgJadwalKerja.Visible = True
        Me.cmdRefresh.Visible = True
    Else
        Me.fgJadwalKerja.Visible = False
        Me.fgJadwalKerja.clear
        Me.dcShiftKerja.Enabled = True
        Me.cmdRefresh.Visible = False
        Call subdgJadwalKerja
    End If
End Sub

Private Sub cmdOtomatis_Click()
    With frmAutoJadwal
        .dcTempatBertugas.Text = Me.dcTempatBertugas.Text
        .DTPicker1.Value = Me.DTPicker1.Value
        If .lvjadwalkerja.ListItems.Count > 0 Then
            .lvjadwalkerja.SetFocus
        End If
        .Show
    End With
End Sub

Private Sub cmdRefresh_Click()
    Call subSetFGJadwalKerja
    Call subIsiFgJadwalKerja
    Me.fgJadwalKerja.SetFocus
End Sub

Private Sub dgJadwalKerja_Click()
'    MsgBox Me.dgJadwalKerja.Row & "/" & Me.dgJadwalKerja.ApproxCount & "," & Me.dgJadwalKerja.Columns("Nama").Value
End Sub

Private Sub DTPicker1_Change()
'    Call subLoadGridTanggal
    intRowCur = 0
    intColCur = 0
    intRowLama = 0
    intColLama = 0
    intTglLibur = 0
    blnHariLibur = False
    Me.DTPicker1.Day = 1
    Call subLoadKalender
    If Me.fgKalender.Rows = 1 Then
        Me.chkTabelJadwal.Enabled = False
        Exit Sub
    Else
        Me.chkTabelJadwal.Enabled = True
    End If
    If Me.fgJadwalKerja.Visible Then
'        Me.chkTabelJadwal.Value = 0
        Call subSetFGJadwalKerja
        Call subIsiFgJadwalKerja
    End If
End Sub

Private Sub fgJadwalKerja_DblClick()
    Call fgJadwalKerja_KeyPress(13)
End Sub

Private Sub fgJadwalKerja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With Me.fgJadwalKerja
            .TextMatrix(.Row, .Col) = ""
        End With
    End If
End Sub

Private Sub fgJadwalKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Me.fgJadwalKerja
            If .TextMatrix(.Row, 0) = "" Then Exit Sub
            If .TextMatrix(1, .Col) = "" Then Exit Sub
            Select Case .TextMatrix(.Row, .Col)
                Case ""
                    .TextMatrix(.Row, .Col) = "P"
                Case "P"
                    .TextMatrix(.Row, .Col) = "S"
                Case "S"
                    .TextMatrix(.Row, .Col) = "M"
                Case "M"
                    .TextMatrix(.Row, .Col) = "L"
                Case "L"
                    .TextMatrix(.Row, .Col) = ""
            End Select
        End With
    End If
End Sub

Private Sub fgJadwalKerja_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Button = 1
        PopupMenu MDIUtama.mnuPopUp
    End If
End Sub

Private Sub fgKalender_Click()
    On Error GoTo tangani
    
    With Me.fgKalender
        If .TextMatrix(.Row, .Col) = "" Then Exit Sub
        Me.DTPicker1.Day = .TextMatrix(.Row, .Col)
        strSQL = "SELECT [Kode Tgl], Tanggal, [Hari Libur] FROM v_tanggal" & _
                 " WHERE DAY(Tanggal)='" & Day(Me.DTPicker1.Value) & "'" & _
                 " AND MONTH(Tanggal)='" & Month(Me.DTPicker1.Value) & "'" & _
                 " AND YEAR(Tanggal)='" & Year(Me.DTPicker1.Value) & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount > 0 Then
            Me.txtkdtgl.Text = rs.Fields.Item("Kode Tgl").Value
            Me.txtnamatgl.Text = rs.Fields.Item("Tanggal").Value
            Me.txtharilibur.Text = IIf(IsNull(rs.Fields.Item("Hari Libur").Value), "", rs.Fields.Item("Hari Libur").Value)
            Me.lblKeterangan.Caption = IIf(Me.txtharilibur.Text <> "", FormatDateTime(Me.DTPicker1.Value, vbLongDate) & ": " & Me.txtharilibur.Text, "")
        End If
'        If .CellBackColor = vbRed Then blnHariLibur = True
        If Me.txtharilibur.Text <> "" Then intTglLibur = Day(Me.txtnamatgl.Text)
        .CellBackColor = &H808080
        intRowCur = .Row
        intColCur = .Col
        If intRowLama = 0 Then GoTo jump
        If intRowCur = intRowLama And intColCur = intColLama Then GoTo jump
        .Row = intRowLama
        .Col = intColLama
        If .TextMatrix(.Row, .Col) = Day(Now) And Me.DTPicker1.Month = Month(Now) Then
            .CellBackColor = &H4000&
        ElseIf .TextMatrix(.Row, .Col) = intTglLibur Then
            .CellBackColor = vbRed
            blnHariLibur = False
        Else
            .CellBackColor = vbWhite
        End If
jump:
        intRowLama = intRowCur
        intColLama = intColCur
        .Row = intRowLama
        .Col = intColLama
    End With
    If Me.fgJadwalKerja.Visible = True Then
        Call subTandaiTanggal(Me.DTPicker1.Value)
        Exit Sub
    End If
    Call loadListViewSource
    Exit Sub
    
tangani:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
'        With frmLaporanJadwalKerja
'            .Show
'            .dcRuangan.Text = Me.dcTempatBertugas.Text
'            .dtpAwal.Year = Me.DTPicker1.Year
'            .dtpAwal.Month = Me.DTPicker1.Month
'            .dtpAwal.Day = 1
'            .dtpAkhir.Year = Me.DTPicker1.Year
'            .dtpAkhir.Month = Me.DTPicker1.Month
'            .dtpAkhir.Day = Me.DTPicker1.Day
'        End With
    End If
End Sub

Public Sub loadListViewSource()
    On Error GoTo tangani

    strSQL = "SELECT IdPegawai, NamaLengkap From v_TempatBertugas WHERE NamaRuangan = '" & dcTempatBertugas & "' ORDER BY IdPegawai"
    Call msubRecFO(rs, strSQL)
    lvjadwalkerja.ListItems.clear
    While Not rs.EOF
        lvjadwalkerja.ListItems.add , "A" & rs(0).Value, rs(1).Value
        rs.MoveNext
    Wend
    lvjadwalkerja.Sorted = True

    If rs.RecordCount = 0 Then Exit Sub
    strSQL = "SELECT ID from v_JadwalKerja WHERE KdTgl = '" & txtkdtgl.Text & "' AND KdShift = '" & dcShiftKerja.BoundText & "' AND KdRuangan = '" & dcTempatBertugas.BoundText & "' "
    Call msubRecFO(rs, strSQL)
    Do While rs.EOF = False
        lvjadwalkerja.ListItems("A" & rs(0)).Checked = True
        lvjadwalkerja.ListItems("A" & rs(0)).ForeColor = vbBlue
        lvjadwalkerja.ListItems("A" & rs(0)).Bold = True
        rs.MoveNext
    Loop
    Exit Sub

tangani:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Sub subDcSource()
   strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan order by NamaRuangan"
   Call msubDcSource(dcTempatBertugas, rs, strSQL)
        
   strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja order by NamaShift"
   Call msubDcSource(dcShiftKerja, rs, strSQL)
End Sub

Private Sub cmdCetak_Click()
frmCetakJadwalKerja.Show
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad
    If Me.fgJadwalKerja.Visible = True Then
        Call subSimpanJadwalBaru
        Exit Sub
    End If
    If Periksa("text", txtkdtgl, "Data Tanggal Belum di Pilih") = False Then Exit Sub
    If Periksa("datacombo", dcShiftKerja, "Shift Kerja Belum Diisi") = False Then Exit Sub
    If Periksa("text", txtkdtgl, "Kode Tanggal Belum Dipilih") = False Then Exit Sub
    
    For i = 1 To lvjadwalkerja.ListItems.Count
       ' strSQL = "DELETE jadwalkerja WHERE KdTgl = '" & txtkdtgl.Text & "' AND IdPegawai = '" & lvjadwalkerja.ListItems(i).key & "'"
       ' dbConn.Execute strSQL
        If lvjadwalkerja.ListItems(i).Checked = True Then
            If sp_JadwalKerja(Right(lvjadwalkerja.ListItems(i).key, Len(lvjadwalkerja.ListItems(i).key) - 1)) = False Then Exit Sub
        Else
           strSQL = "DELETE JadwalKerja WHERE KdShift = '" & dcShiftKerja.BoundText & "' AND KdTgl = '" & txtkdtgl.Text & "' AND IdPegawai = '" & Right(lvjadwalkerja.ListItems(i).key, Len(lvjadwalkerja.ListItems(i).key) - 1) & "'"
            dbConn.Execute strSQL
        End If
    Next i
    'MsgBox "Data Berhasil Disimpan", vbInformation, "Informasi"
    Call loadListViewSource
    Call subdgJadwalKerja
    Exit Sub
errLoad:
    Call msubPesanError
End Sub
Public Function sp_JadwalKerja(f_KdPegawai As String) As Boolean
    sp_JadwalKerja = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdTgl", adVarChar, adParamInput, 3, txtkdtgl.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_KdPegawai)
        .Parameters.Append .CreateParameter("KdShift", adChar, adParamInput, 2, dcShiftKerja.BoundText)
        '.Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AU_JADWALKERJA"
        .Execute
        
'        If Not (.Parameters("RETURN_VALUE").Value) = 0 Then
'            MsgBox "Ada kesalahan dalam pemasukan data", vbCritical, "Validasi"
'            sp_JadwalKerja = False
'        End If
        
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

'Private Sub subLoadGridTanggal()
'On Error GoTo errLoad
'strSQL = "SELECT [Kode Tgl], Tanggal, [Hari Libur] FROM v_tanggal" & _
'         " WHERE MONTH(Tanggal)='" & Month(Me.DTPicker1.Value) & "'" & _
'         " AND YEAR(Tanggal)='" & Year(Me.DTPicker1.Value) & "'"
'Call msubRecFO(rs, strSQL)
'Set dgTanggal.DataSource = rs
'With dgTanggal
'    .Columns("Kode Tgl").Width = 0
'    .Columns("Tanggal").Width = 1000
'    .Columns("Hari Libur").Width = 2000
'    .Visible = True
'End With
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub

Private Sub dcShiftKerja_Change()
    Call loadListViewSource
    Call subdgJadwalKerja
End Sub

Private Sub dcTempatBertugas_Change()
Call loadListViewSource
Call subdgJadwalKerja
If Me.lvjadwalkerja.ListItems.Count > 0 Then
'    If Me.fgKalender.Rows = 1 Then
'        Me.chkTabelJadwal.Enabled = False
'        Exit Sub
'    Else
        Me.chkTabelJadwal.Enabled = True
'    End If
'    If Me.chkTabelJadwal.Value = 1 Then
'        Call subSetFGJadwalKerja
'        Call subIsiFgJadwalKerja
'    End If
Else
'    Me.chkTabelJadwal.Enabled = False
'    If Me.chkTabelJadwal.Value = 1 Then Me.chkTabelJadwal.Value = 0
End If
End Sub

Private Sub dcTempatBertugas_GotFocus()
On Error GoTo errLoad
    strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan Order By NamaRuangan"
    Call msubDcSource(dcTempatBertugas, rs, strSQL)
Exit Sub
errLoad:
End Sub

Private Sub dgtanggal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
With dgTanggal
     txtkdtgl.Text = .Columns("Kode Tgl").Value
     txtnamatgl.Text = .Columns("Tanggal").Value
     If .Columns("Hari Libur").Value = "" Then
        txtharilibur.Text = ""
     Else
        txtharilibur.Text = .Columns("Hari Libur").Value
     End If
Call loadListViewSource
End With
End Sub

Private Sub Form_Load()
'    Me.mnuPopUpMenu.Visible = False
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    Me.DTPicker1.Value = Now
    Call subDcSource
'    Call subLoadGridTanggal
    Call loadListViewSource
    Call subdgJadwalKerja
'    Call subSetFgKalender
'    Call subLoadKalender
End Sub

Private Sub lvjadwalkerja_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim intRowNama As Integer
    Dim intMaxCol As Integer, intMaxRow As Integer
    Dim i As Integer, j As Integer
    
    With Me.fgJadwalKerja
        intMaxCol = .Cols - 1
        intMaxRow = .Rows - 1
        
        For i = 2 To intMaxRow
            .Col = 0
            .Row = i
            If .CellBackColor = &H808080 Then
                For j = 0 To intMaxCol
                    .Col = j
                    If .CellBackColor = vbYellow Then
                        .CellBackColor = &H808080
                    Else
                        .CellBackColor = vbWhite
                    End If
                Next
                .Col = 0
                .CellBackColor = &H8000000F
            End If
        Next
        intRowNama = funcCariRowNama(Item.Text)
        .Row = intRowNama
        For i = 0 To intMaxCol
            .Col = i
            If .CellBackColor = &H808080 Then
                .CellBackColor = vbYellow
            Else
                .CellBackColor = &H808080
            End If
        Next
    End With
End Sub

Private Sub txtNamaPegawai_Change()
strSQL = "SELECT NamaLengkap, IdPegawai FROM v_Tempatbertugas where NamaRuangan LIKE '" & dcTempatBertugas & "' AND NamaLengkap LIKE '" & txtnamapegawai & "%' ORDER by NamaLengkap "
Call msubRecFO(rs, strSQL)
    lvjadwalkerja.ListItems.clear
    While Not rs.EOF
        lvjadwalkerja.ListItems.add , "A" & rs(1).Value, rs(0).Value
        rs.MoveNext
    Wend
    lvjadwalkerja.Sorted = True
Exit Sub
End Sub

Private Sub subdgJadwalKerja()
On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "SELECT * FROM V_JadwalKerja WHERE Ruangan='" & dcTempatBertugas & "'" & _
             " ORDER BY Tanggal"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgJadwalKerja.DataSource = rs
    With dgJadwalKerja
        .Columns("KdShift").Width = 0
        .Columns("Nama").Width = 2500
        .Columns("KdRuangan").Width = 0
        .Columns("KdTgl").Width = 0
    End With
Exit Sub
Set rs = Nothing
errLoad:
    Call msubPesanError
End Sub

'Tambahan 16 April 2008
Private Sub subSetFgKalender()
    Dim i As Integer
    Dim intLebarCell As Integer
    
    With Me.fgKalender
        .clear
        .Rows = 1
        .Cols = 7
        
        .RowHeight(0) = 500
        
        .TextMatrix(0, 0) = "Sen"
        .TextMatrix(0, 1) = "Sel"
        .TextMatrix(0, 2) = "Rab"
        .TextMatrix(0, 3) = "Kam"
        .TextMatrix(0, 4) = "Jum"
        .TextMatrix(0, 5) = "Sab"
        .TextMatrix(0, 6) = "Ming"
        .Row = 0
        .Col = 6
        .CellBackColor = vbRed
        intLebarCell = (Me.fgKalender.Width / 7) - 2
        For i = 0 To 6
            .ColWidth(i) = intLebarCell
            .Row = 0
            .Col = i
            .CellFontBold = True
        Next
    End With
End Sub

'Private Sub subLoadKalender()
'    Dim tgl As Date
'    Dim strTanggal As String, strBulan As String, strTahun As String
'    Dim strHari As String
'    Dim intCol As Integer, intRow As Integer
'    Dim blnAwalBulan As Boolean
'
'    Call subSetFgKalender
'
'    strSQL = "SELECT [Kode Tgl], Tanggal, [Hari Libur] FROM v_tanggal" & _
'             " WHERE MONTH(Tanggal)='" & Month(Me.DTPicker1.Value) & "'" & _
'             " AND YEAR(Tanggal)='" & Year(Me.DTPicker1.Value) & "'" & _
'             " ORDER BY Tanggal"
'    Call msubRecFO(rs, strSQL)
'
'    blnAwalBulan = True
'    While Not rs.EOF
'        With Me.fgKalender
'            tgl = rs.Fields.Item("Tanggal").Value
'            strTanggal = Day(tgl)
'            strBulan = Month(tgl)
'            strHari = WeekdayName(Weekday(tgl), , vbSunday)
'
'            If blnAwalBulan Then
'                .Rows = .Rows + 1
'                intRow = .Rows - 1
'                blnAwalBulan = False
'            End If
'            Select Case strHari
'                Case "Senin"
'                    intCol = 0
'                    intRow = .Rows - 1
'                Case "Selasa"
'                    intCol = 1
'                Case "Rabu"
'                    intCol = 2
'                Case "Kamis"
'                    intCol = 3
'                Case "Jumat"
'                    intCol = 4
'                Case "Sabtu"
'                    intCol = 5
'                Case "Minggu"
'                    intCol = 6
'                    .Rows = .Rows + 1
'            End Select
'            .TextMatrix(intRow, intCol) = strTanggal
'            If strTanggal = Day(Now) And strBulan = Month(Now) Then
'                .Row = intRow
'                .Col = intCol
'                .CellBackColor = &H4000&
'                .CellForeColor = vbWhite
'            End If
'            If strHari = "Minggu" Then
'                .Row = intRow
'                .Col = intCol
'                .CellForeColor = vbRed
'            End If
'            If Not IsNull(rs.Fields.Item("Hari Libur").Value) Then
'                .Row = intRow
'                .Col = intCol
'                .CellBackColor = vbRed
'                .CellFontBold = True
'                .CellForeColor = vbWhite
'            End If
'            rs.MoveNext
'        End With
'    Wend
'End Sub

'Tambahan 24 April 2008
Private Sub subSetFGJadwalKerja()
    Dim intJumHari As Integer
    Dim i As Integer
    Dim itm As ListItem
    Dim strHari As String
    
    Call totalHari(Me.DTPicker1.Month, Me.DTPicker1.Year)
    intJumHari = TtlHari
    With Me.fgJadwalKerja
        .clear
        .Rows = 3
        .Cols = intJumHari + 1
        .FixedCols = 1
        .FixedRows = 2
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
        .TextMatrix(0, 0) = "Nama Pegawai"
        .TextMatrix(1, 0) = "Nama Pegawai"
        .ColWidth(0) = 3000
        .RowHeight(0) = 300
        For i = 1 To intJumHari
            Me.DTPicker1.Day = i
            strHari = WeekdayName(Weekday(Me.DTPicker1.Value), , vbSunday)
            
            .TextMatrix(0, i) = Format(Me.DTPicker1.Value, "MMMM yyyy")
            .TextMatrix(1, i) = i
            .Row = 0
            .Col = i
            .CellAlignment = flexAlignCenterCenter
            .ColWidth(i) = 500
            .ColAlignment(i) = flexAlignCenterCenter
            If strHari = "Minggu" Then
                .Row = 1
                .CellBackColor = vbRed
                .CellForeColor = vbWhite
            Else
                .Row = 1
                .CellBackColor = &H4000&
                .CellForeColor = vbWhite
            End If
            strSQL = "SELECT [Hari Libur] FROM v_tanggal" & _
                     " WHERE DAY(Tanggal)='" & i & "'" & _
                     " AND MONTH(Tanggal)='" & Month(Me.DTPicker1.Value) & "'" & _
                     " AND YEAR(Tanggal)='" & Year(Me.DTPicker1.Value) & "'"
            Call msubRecFO(rs, strSQL)
            If Not IsNull(rs.Fields.Item("Hari Libur").Value) Then
                .Row = 1
                .CellBackColor = vbRed
                .CellForeColor = vbWhite
            End If
        Next
        For Each itm In Me.lvjadwalkerja.ListItems
            .Row = .Rows - 1
            .TextMatrix(.Row, 0) = itm.Text
            .Rows = .Rows + 1
        Next
    End With
End Sub

'Tambahan 25 April 2008

Private Sub subIsiFgJadwalKerja()
    Dim intRowNama As Integer, intColTgl As Integer
    Dim strNama As String, strTanggal As String, strShift As String
    Dim strKodeShift As String
    Dim i As Integer
    
'    If Me.dgJadwalKerja.ApproxCount < 0 Then Exit Sub
    
    strSQL = "SELECT * FROM V_JadwalKerja" & _
             " WHERE Ruangan='" & dcTempatBertugas & "'" & _
             " AND MONTH(Tanggal)='" & Me.DTPicker1.Month & "'" & _
             " AND YEAR(Tanggal)='" & Me.DTPicker1.Year & "'" & _
             " ORDER BY Tanggal"
    Call msubRecFO(rs, strSQL)
    While Not rs.EOF
        strNama = rs.Fields.Item("Nama").Value
        strShift = rs.Fields.Item("Shift").Value
        strTanggal = rs.Fields.Item("Tanggal").Value
        
        intRowNama = funcCariRowNama(strNama)
        intColTgl = funcCariColTanggal(Day(strTanggal))
        
        strKodeShift = UCase$(Left(strShift, 1))
        Me.fgJadwalKerja.TextMatrix(intRowNama, intColTgl) = strKodeShift
        rs.MoveNext
    Wend
End Sub

Private Function funcCariRowNama(ByVal NamaPegawai As String) As Integer
    Dim intMaxRow As Integer
    Dim i As Integer
    
    With Me.fgJadwalKerja
        intMaxRow = .Rows - 1
        For i = 2 To intMaxRow
            If .TextMatrix(i, 0) = NamaPegawai Then
                funcCariRowNama = i
                Exit For
            Else
                funcCariRowNama = 0
            End If
        Next
    End With
End Function

Private Function funcCariColTanggal(ByVal tgl As String) As Integer
    Dim intMaxCol As Integer
    Dim i As Integer
    
    With Me.fgJadwalKerja
        intMaxCol = .Cols - 1
        For i = 1 To intMaxCol
            If .TextMatrix(1, i) = tgl Then
                funcCariColTanggal = i
                Exit For
            Else
                funcCariColTanggal = 0
            End If
        Next
    End With
End Function

'Tambahan 28 April 2008
Private Sub subSimpanJadwalBaru()
    On Error GoTo tangani
    
    Dim intMaxRow As Integer, intMaxCol As Integer
    Dim c As Integer, R As Integer
    Dim strKdTgl As String, strKdShift As String, strIDPegawai As String
    
    With Me.fgJadwalKerja
        intMaxRow = .Rows - 1
        intMaxCol = .Cols - 1
        
        For R = 2 To intMaxRow
            strIDPegawai = funcCariIdPegawai(.TextMatrix(R, 0))
            For c = 1 To intMaxCol
                strKdShift = funcCariKodeShift(.TextMatrix(R, c))
                Me.DTPicker1.Day = CInt(.TextMatrix(1, c))
                strSQL = "SELECT [Kode Tgl] FROM v_tanggal" & _
                         " WHERE DAY(Tanggal)='" & Me.DTPicker1.Day & "'" & _
                         " AND MONTH(Tanggal)='" & Me.DTPicker1.Month & "'" & _
                         " AND YEAR(Tanggal)='" & Me.DTPicker1.Year & "'"
                Call msubRecFO(rs, strSQL)
                If Not rs.EOF Then
                    strKdTgl = rs.Fields.Item("Kode Tgl").Value
                End If
                If .TextMatrix(R, c) = "" Then
                    Call subDeleteJadwalPerShift(strIDPegawai, strKdTgl)
                Else
                    Call subJadwalKerja(strKdTgl, strIDPegawai, strKdShift)
                End If
            Next
        Next
    End With
    MsgBox "Simpan jadwal selesai.", vbInformation, "Perhatian"
    Exit Sub
    
tangani:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Function funcCariIdPegawai(ByVal NamaPegawai As String) As String
    Dim itm As ListItem
    
    For Each itm In Me.lvjadwalkerja.ListItems
        If itm.Text = NamaPegawai Then
            funcCariIdPegawai = Right(itm.key, Len(itm.key) - 1)
            Exit For
        Else
            funcCariIdPegawai = ""
        End If
    Next
End Function

Private Function funcCariKodeShift(ByVal NamaShift As String) As String
    strSQL = "SELECT KdShift FROM ShiftKerja" & _
             " WHERE NamaShift LIKE '" & NamaShift & "%'"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        funcCariKodeShift = rs.Fields.Item("KdShift").Value
    End If
End Function

Public Sub subJadwalKerja(ByVal KdTgl As String, ByVal idpegawai As String, ByVal KdShift As String)
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdTgl", adVarChar, adParamInput, 3, KdTgl)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, idpegawai)
        .Parameters.Append .CreateParameter("KdShift", adChar, adParamInput, 2, KdShift)
        
        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AU_JADWALKERJA"
        .Execute
  
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Sub

Private Sub subTandaiTanggal(ByVal TanggalPilihan As String)
    Dim strTgl As String
    Dim i As Integer, j As Integer
    Dim intMaxRow As Integer, intMaxCol As Integer
    
    strTgl = Day(TanggalPilihan)
    With Me.fgJadwalKerja
        intMaxRow = .Rows - 1
        intMaxCol = .Cols - 1
        For i = 1 To intMaxCol
            .Row = 1
            .Col = i
            If .CellBackColor = &H808080 Then
                For j = 1 To intMaxRow
                    .Row = j
                    If .CellBackColor = vbYellow Then
                        .CellBackColor = &H808080
                    Else
                        .CellBackColor = vbWhite
                    End If
                Next
                .Row = 1
                Me.DTPicker1.Day = CInt(.TextMatrix(.Row, .Col))
                If WeekdayName(Weekday(Me.DTPicker1.Value), , vbSunday) = "Minggu" Then
                    .CellBackColor = vbRed
                Else
                    .CellBackColor = &H4000&
                End If
            End If
            If strTgl = .TextMatrix(1, i) Then
                For j = 1 To intMaxRow
                    .Row = j
                    If .CellBackColor = &H808080 Then
                        .CellBackColor = vbYellow
                    Else
                        .CellBackColor = &H808080
                    End If
                Next
'                Exit For
            End If
        Next
    End With
End Sub

Private Sub subDeleteJadwalPerShift(ByVal idpegawai As String, ByVal KodeTanggal As String)
    strSQL = "DELETE FROM JadwalKerja WHERE" & _
             " KdTgl='" & KodeTanggal & "'" & _
             " AND IdPegawai='" & idpegawai & "'"
    dbConn.Execute strSQL
End Sub

