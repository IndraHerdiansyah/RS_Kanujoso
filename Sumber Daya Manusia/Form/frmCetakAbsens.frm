VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Begin VB.Form frmAbsensiPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Absensi Pegawai "
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   Icon            =   "frmCetakAbsens.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11325
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   7740
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   11360
            MinWidth        =   11360
            Text            =   "F1 - Cetak Daftar Hadir Karyawan/ti "
            TextSave        =   "F1 - Cetak Daftar Hadir Karyawan/ti "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   11359
            MinWidth        =   11359
            Text            =   "F2 - Cetak Detail Kehadiran Karyawan/ti "
            TextSave        =   "F2 - Cetak Detail Kehadiran Karyawan/ti "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periode "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   9375
      Begin VB.OptionButton optJamKeluar 
         Caption         =   "Dinas Pulang"
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optJamMasuk 
         Caption         =   "Dinas Datang"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         Height          =   330
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpTglAhkir 
         Height          =   360
         Left            =   6120
         TabIndex        =   6
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy HH:mm"
         Format          =   59375619
         UpDown          =   -1  'True
         CurrentDate     =   38231
      End
      Begin MSComCtl2.DTPicker dtpTglAwal 
         Height          =   360
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy HH:mm"
         Format          =   59375619
         UpDown          =   -1  'True
         CurrentDate     =   38231
      End
      Begin VB.Label Label1 
         Caption         =   "s/d"
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   11295
      Begin VB.TextBox txtCariPegawai 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9600
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Masukan Nama Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSDataGridLib.DataGrid dgAbsensi 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8493
      _Version        =   393216
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1800
      _cx             =   4197479
      _cy             =   4196024
      FlashVars       =   ""
      Movie           =   "Window"
      Src             =   "Window"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin MSDataListLib.DataCombo dcJenisPegawai 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Pegawai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmCetakAbsens.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9480
      Picture         =   "frmCetakAbsens.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmCetakAbsens.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmAbsensiPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCommand As New ADODB.Command
Dim mdtBulan As Integer
Dim MdtTahun As Integer
Private Sub cmdCari_Click()
    Call subLoadGridSource
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisPegawai.MatchedWithList = True Then dtpTglAwal.SetFocus
        strSQL = "SELECT KdJenisPegawai, JenisPegawai FROM JenisPegawai  where JenisPegawai like '%" & dcJenisPegawai.Text & "%' Order By JenisPegawai"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcJenisPegawai.BoundText = rs(0).Value
        dcJenisPegawai.Text = rs(1).Value
    End If
End Sub

Private Sub dgAbsensi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgAbsensi.ApproxCount = 0 Then Exit Sub
    txtNoAbsensi.Text = dgAbsensi.Columns("NoAbsensi")
    dcNamaPegawai.Text = dgAbsensi.Columns("NamaPegawai")
    dcStatusMasuk.Text = dgAbsensi.Columns("StatusMasuk")
    dtpMasuk.Value = dgAbsensi.Columns("TglMasuk")
    dtpPulang.Value = dgAbsensi.Columns("TglPulang")
    
    cmdhapus.Enabled = True
End Sub



Private Sub dtpTglAhkir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpTglAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglAhkir.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            If Len(Trim(dcJenisPegawai.Text)) = 0 Then MsgBox "Isi Jenis Pegawai!!", vbInformation, "Informasi": Exit Sub
            mstrCetak2 = dcJenisPegawai.Text
            strCetak = "CetakAbsensi"
            mdTglAwal = dtpTglAwal.Value
            mdTglAkhir = dtpTglAhkir.Value
            frmCetakAbsensiPegawai.Show
        Case vbKeyF2
            If Len(Trim(dcJenisPegawai.Text)) = 0 Then MsgBox "Isi Jenis Pegawai!!", vbInformation, "Informasi": Exit Sub
            If optJamMasuk.Value = True Then mstrGroup = "1" Else mstrGroup = "2"
            
            mstrCetak2 = dcJenisPegawai.Text
            strCetak = "DetailCetakAbsensi"
            mdTglAwal = dtpTglAwal.Value
            mdTglAkhir = dtpTglAhkir.Value
            frmCetakAbsensiPegawai.Show
    End Select

End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglAwal.Value = Format(Now, " dd/MMMM/yyyy 00:00:00")
    dtpTglAhkir.Value = Format(Now, " dd/MMMM/yyyy 23:59:59")
    optJamMasuk.Value = True
    Call subLoadDcSource
    Call subLoadGridSource
End Sub
Private Sub subLoadDcSource()
    Call msubDcSource(dcJenisPegawai, rs, "SELECT KdJenisPegawai, JenisPegawai FROM JenisPegawai Order By JenisPegawai")
End Sub
'
'
Sub subLoadGridSource()
On Error GoTo errSimpan
        'mdtBulan = CStr(Format(dtpTglAhkir.Value, "mm"))
        'MdtTahun = CStr(Format(dtpTglAhkir.Value, "yyyy"))
        'mdTglAkhir = CDate(Format(dtpTglAhkir.Value, "yyyy-mm") & "-" & funcHitungHari(mdtBulan, MdtTahun) & " 23:59:59")

If optJamMasuk.Value = True Then
    Set rs = Nothing
    strSQL = "SELECT TOP 100 PERCENT IdPegawai, PR_FingerID, NamaLengkap, NamaJabatan, NIP, PR_Time" & _
            " From dbo.[V_AbsensiKaryawan/ti] Where (PR_Presence = 1) AND JenisPegawai like '%" & dcJenisPegawai.Text & "%' AND " & _
            " NamaLengkap like '%" & txtCariPegawai.Text & "%' and " & _
            " PR_Time between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd HH:mm:00") & "'AND '" & Format(dtpTglAhkir.Value, "yyyy/MM/dd HH:mm:59") & "'" & _
            " ORDER BY NamaLengkap"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgAbsensi.DataSource = rs
    Set rs = Nothing
 ElseIf optJamKeluar.Value = True Then
 Set rs = Nothing
    strSQL = "SELECT TOP 100 PERCENT IdPegawai, PR_FingerID, NamaLengkap, NamaJabatan, NIP, PR_Time" & _
            " From dbo.[V_AbsensiKaryawan/ti] Where (PR_Presence = 2) AND JenisPegawai like '%" & dcJenisPegawai.Text & "%' AND " & _
            " NamaLengkap like '%" & txtCariPegawai.Text & "%' and " & _
            " PR_Time between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd HH:mm:00") & "'AND '" & Format(dtpTglAhkir.Value, "yyyy/MM/dd HH:mm:59") & "'" & _
            " ORDER BY NamaLengkap"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgAbsensi.DataSource = rs
    Set rs = Nothing
 End If
        dgAbsensi.Columns("PR_Time").Width = 2100
        If optJamMasuk.Value = True Then
            dgAbsensi.Columns("PR_Time").Caption = "Waktu Absensi Masuk"
        ElseIf optJamKeluar.Value = True Then
            dgAbsensi.Columns("PR_Time").Caption = "Waktu Absensi Keluar"
        End If
        dgAbsensi.Columns("IdPegawai").Width = 1400
        dgAbsensi.Columns("IdPegawai").Caption = "ID Pegawai"
        dgAbsensi.Columns("PR_FingerID").Width = 800
        dgAbsensi.Columns("PR_FingerID").Alignment = vbCenter
        dgAbsensi.Columns("PR_FingerID").Caption = "No Absensi"
        dgAbsensi.Columns("NamaLengkap").Width = 2500
        dgAbsensi.Columns("NamaLengkap").Caption = "Nama Pegawai"
        dgAbsensi.Columns("NIP").Width = 1200
        dgAbsensi.Columns("NamaJabatan").Width = 2500
        dgAbsensi.Columns("NamaJabatan").Caption = "Jabatan"
Exit Sub
errSimpan:
    msubPesanError
End Sub

Sub clear()
    txtNoAbsensi.Text = ""
    dcNamaPegawai.Text = ""
    dcStatusMasuk.Text = ""
    txtOvertime.Text = ""
    
    dtpMasuk.Value = Format(Now)
    dtpPulang.Value = Format(Now)
End Sub


Private Sub optBulan_Click()
End Sub

Private Sub optHari_Click()
End Sub

Private Sub optJamKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisPegawai.SetFocus
End Sub

Private Sub optJamMasuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisPegawai.SetFocus
End Sub

Private Sub txtCariPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari_Click
End Sub
