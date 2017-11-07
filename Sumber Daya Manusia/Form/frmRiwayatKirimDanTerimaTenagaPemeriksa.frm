VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatKirimDanTerimaTenagaPemeriksa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Kirim & Terima Tenaga Pemeriksa"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRiwayatKirimDanTerimaTenagaPemeriksa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   12510
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   600
      Left            =   4800
      TabIndex        =   11
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txtNoRiwayat 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   0
      MaxLength       =   30
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   600
      Left            =   6360
      TabIndex        =   12
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   600
      Left            =   7920
      TabIndex        =   13
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   600
      Left            =   11040
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   600
      Left            =   9480
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Kirim/Terima"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   18
      Top             =   1920
      Width           =   12495
      Begin VB.TextBox txtAlamatRujukan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   8
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtTmptRujukan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtKetLainnya 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1800
         Width           =   10335
      End
      Begin MSDataListLib.DataCombo dcDokter 
         Height          =   330
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcRujukanAsal 
         Height          =   330
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcKualifikasiJurusan 
         Height          =   330
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcSubInstalasi 
         Height          =   330
         Left            =   2040
         TabIndex        =   5
         Top             =   1440
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSComCtl2.DTPicker dtpTglKirimTerima 
         Height          =   330
         Left            =   9000
         TabIndex        =   9
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy HH:mm "
         Format          =   108068867
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpTglKembali 
         Height          =   330
         Left            =   9000
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy HH:mm "
         Format          =   108068867
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan Lainnya"
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Kembali"
         Height          =   210
         Index           =   7
         Left            =   5760
         TabIndex        =   26
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Kirim/Terima"
         Height          =   210
         Index           =   6
         Left            =   5760
         TabIndex        =   25
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Tempat Rujukan Tujuan/Asal"
         Height          =   210
         Index           =   5
         Left            =   5760
         TabIndex        =   24
         Top             =   720
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Rujukan Tujuan/Asal"
         Height          =   210
         Index           =   4
         Left            =   5760
         TabIndex        =   23
         Top             =   360
         Width           =   2370
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kasus Penyakit"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kualifikasi Jurusan"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rujukan Asal/Tujuan"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Dokter"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kirim/Terima Tenaga Pemeriksa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   17
      Top             =   1080
      Width           =   12495
      Begin VB.OptionButton optTerima 
         Caption         =   "Riwayat Terima Tenaga Pemeriksa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
      Begin VB.OptionButton optKirim 
         Caption         =   "Riwayat Kirim Tenaga Pemeriksa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   3615
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
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
   End
   Begin MSDataGridLib.DataGrid dgData 
      Height          =   2535
      Left            =   0
      TabIndex        =   28
      Top             =   4320
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatKirimDanTerimaTenagaPemeriksa.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10680
      Picture         =   "frmRiwayatKirimDanTerimaTenagaPemeriksa.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatKirimDanTerimaTenagaPemeriksa.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmRiwayatKirimDanTerimaTenagaPemeriksa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub subKosong()
    On Error GoTo hell
    dcDokter.Text = ""
    dcRujukanAsal.Text = ""
    dcKualifikasiJurusan.Text = ""
    dcSubInstalasi.Text = ""
    txtKetLainnya.Text = ""
    txtTmptRujukan.Text = ""
    txtAlamatRujukan.Text = ""
    dtpTglKirimTerima.Value = Now
    dtpTglKembali.Value = Now
    txtnoriwayat.Text = ""
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadDcSource()
    On Error GoTo hell
    'Call msubDcSource(dcDokter, rs, "Select IdPegawai,NamaLengkap From DataPegawai")
    Call msubDcSource(dcDokter, rs, "Select IdPegawai,[Nama Lengkap] from V_M_DataPegawaiNew where KdJenisPegawai='001' and KdStatus='01'")
    Call msubDcSource(dcRujukanAsal, rs, "Select KdRujukanAsal,RujukanAsal From RujukanAsal Where StatusEnabled=1")
    Call msubDcSource(dcKualifikasiJurusan, rs, "Select KdKualifikasiJurusan,KualifikasiJurusan From KualifikasiJurusan Where StatusEnabled=1")
    Call msubDcSource(dcSubInstalasi, rs, "Select KdSubInstalasi,NamaSubInstalasi From SubInstalasi Where StatusEnabled=1")
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadDataGrid()
    On Error GoTo hell
    If optKirim.Value = True Then
        Frame1.Caption = "Kirim Tenaga Pemeriksa"
        Frame2.Caption = "Data Tenaga Pemeriksa - Kirim"
        Label1(0).Caption = "Rujukan Tujuan"
        Label1(4).Caption = "Tempat Rujukan Tujuan"
        Label1(5).Caption = "Alamat Tempat Rujukan Tujuan"
        Label1(6).Caption = "Tgl Kirim"

        strSQL = "Select * From V_RiwayatKirimTenagaPemeriksa"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)
        Set dgData.DataSource = rs
        With dgData
            For i = 1 To .Col
                .Columns(i).Width = 0
            Next i

            .Columns("NoRiwayat").Width = 1000
            .Columns("NamaDokter").Width = 1200
            .Columns("RujukanTujuan").Width = 1000
            .Columns("KualifikasiJurusan").Width = 1000
            .Columns("SubInstalasi").Width = 1300
            .Columns("TempatRujukanTujuan").Width = 1500
            .Columns("AlamatTempatRujukanTujuan").Width = 1500
            .Columns("TglKirim").Width = 1250
            .Columns("TglKembali").Width = 1250
            .Columns("Keterangan").Width = 1500

        End With
    ElseIf optTerima.Value = True Then
        Frame1.Caption = "Terima Tenaga Pemeriksa"
        Frame2.Caption = "Data Tenaga Pemeriksa - Terima"
        Label1(0).Caption = "Rujukan Asal"
        Label1(4).Caption = "Tempat Rujukan Asal"
        Label1(5).Caption = "Alamat Tempat Rujukan Asal"
        Label1(6).Caption = "Tgl Terima"

        strSQL = "Select * From V_RiwayatTerimaTenagaPemeriksa"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)
        Set dgData.DataSource = rs
        With dgData
            For i = 1 To .Col
                .Columns(i).Width = 0
            Next i

            .Columns("NoRiwayat").Width = 1000
            .Columns("NamaDokter").Width = 1200
            .Columns("RujukanAsal").Width = 1000
            .Columns("KualifikasiJurusan").Width = 1000
            .Columns("SubInstalasi").Width = 1300
            .Columns("TempatRujukanAsal").Width = 1500
            .Columns("AlamatTempatRujukanAsal").Width = 1500
            .Columns("TglTerima").Width = 1250
            .Columns("TglKembali").Width = 1250
            .Columns("Keterangan").Width = 1500

        End With
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadDcSource
    Call subLoadDataGrid
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If optKirim.Value = True Then
        If txtnoriwayat.Text = "" Then Exit Sub
        If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
        'strSQL = "DELETE FROM RiwayatKirimTenagaPemeriksa WHERE NoRiwayat='" & txtNoRiwayat.Text & "' "
        'Call msubRecFO(rs, strSQL)
        If sp_RiwayatKirimTenagaPemeriksa2("D") = False Then Exit Sub '//yayang.agus 2014-08-12
'        strSQL = "DELETE FROM Riwayat WHERE NoRiwayat='" & txtNoRiwayat.Text & "' "
'        Call msubRecFO(rs, strSQL)
        If sp_Riwayat("D") = False Then Exit Sub

    Else
        If txtnoriwayat.Text = "" Then Exit Sub
        If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
        'strSQL = "DELETE FROM RiwayatTerimaTenagaPemeriksa WHERE NoRiwayat='" & txtNoRiwayat.Text & "' "
        'Call msubRecFO(rs, strSQL)
        If sp_RiwayatTerimaTenagaPemeriksa2("D") = False Then Exit Sub '//yayang.agus 2014-08-12
'        strSQL = "DELETE FROM Riwayat WHERE NoRiwayat='" & txtNoRiwayat.Text & "' "
'        Call msubRecFO(rs, strSQL)
        If sp_Riwayat("D") = False Then Exit Sub
    End If
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If optKirim.Value = True Then
        If Periksa("datacombo", dcDokter, "Silahkan pilih nama dokter ") = False Then Exit Sub
        If Periksa("datacombo", dcRujukanAsal, "Silahkan pilih Rujukan asal ") = False Then Exit Sub
        If Periksa("datacombo", dcKualifikasiJurusan, "Silahkan pilih Kualifikasi Jurusan ") = False Then Exit Sub
        If Periksa("datacombo", dcSubInstalasi, "Silahkan pilih Kasus penyakit ") = False Then Exit Sub
        If Periksa("text", txtTmptRujukan, "Silahkan isi Tempat rujukan tujuan ") = False Then Exit Sub

'        If sp_RiwayatKirimTenagaPemeriksa("A") = False Then Exit Sub
        Call sp_Riwayat("A") '= False Then Exit Sub
        'Call sp_RiwayatKirimTenagaPemeriksa   '= False Then Exit Sub
        Call sp_RiwayatKirimTenagaPemeriksa2("A")   '//yayang.agus 2014-08-12

    Else
        If Periksa("datacombo", dcDokter, "Silahkan pilih nama dokter ") = False Then Exit Sub
        If Periksa("datacombo", dcRujukanAsal, "Silahkan pilih Rujukan Tujuan ") = False Then Exit Sub
        If Periksa("datacombo", dcKualifikasiJurusan, "Silahkan pilih Kualifikasi Jurusan ") = False Then Exit Sub
        If Periksa("datacombo", dcSubInstalasi, "Silahkan pilih Kasus penyakit ") = False Then Exit Sub
        If Periksa("text", txtTmptRujukan, "Silahkan isi Tempat rujukan Asal ") = False Then Exit Sub

        Call sp_Riwayat("A")
'        Call sp_RiwayatTerimaTenagaPemeriksa '= False Then Exit Sub
        Call sp_RiwayatTerimaTenagaPemeriksa2("A")   '//yayang.agus 2014-08-12
    End If
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Call cmdBatal_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcDokter.MatchedWithList = True Then dcRujukanAsal.SetFocus
        strSQL = "Select IdPegawai,NamaLengkap From DataPegawai WHERE (NamaLengkap LIKE '%" & dcDokter.Text & "%') "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcDokter.BoundText = rs(0).Value
        dcDokter.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKualifikasiJurusan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcKualifikasiJurusan.MatchedWithList = True Then dcSubInstalasi.SetFocus
        strSQL = "Select KdKualifikasiJurusan,KualifikasiJurusan From KualifikasiJurusan WHERE (KualifikasiJurusan LIKE '%" & dcKualifikasiJurusan.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKualifikasiJurusan.BoundText = rs(0).Value
        dcKualifikasiJurusan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcRujukanAsal_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcRujukanAsal.MatchedWithList = True Then dcKualifikasiJurusan.SetFocus
        strSQL = "Select KdRujukanAsal,RujukanAsal From RujukanAsal WHERE (RujukanAsal LIKE '%" & dcRujukanAsal.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRujukanAsal.BoundText = rs(0).Value
        dcRujukanAsal.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcSubInstalasi_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcSubInstalasi.MatchedWithList = True Then txtKetLainnya.SetFocus
        strSQL = "Select KdSubInstalasi,NamaSubInstalasi From SubInstalasi WHERE (NamaSubInstalasi LIKE '%" & dcSubInstalasi.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcSubInstalasi.BoundText = rs(0).Value
        dcSubInstalasi.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo hell
    With dgData
        If .ApproxCount <= 0 Then Exit Sub
        If optKirim.Value = True Then
            txtnoriwayat.Text = .Columns("NoRiwayat")
            dcDokter.BoundText = .Columns("IdDokter")
            dcRujukanAsal.BoundText = .Columns("KdRujukanTujuan")
            dcKualifikasiJurusan.BoundText = .Columns("KdKualifikasiJurusan")
            dcSubInstalasi.BoundText = .Columns("KdSubInstalasi")
            txtTmptRujukan.Text = .Columns("TempatRujukanTujuan")
            txtAlamatRujukan.Text = .Columns("AlamatTempatRujukanTujuan")
            dtpTglKirimTerima.Value = .Columns("TglKirim")
            dtpTglKembali.Value = .Columns("TglKembali")
            txtKetLainnya.Text = .Columns("Keterangan")
        ElseIf optTerima.Value = True Then
            txtnoriwayat.Text = .Columns("NoRiwayat")
            dcDokter.BoundText = .Columns("IdDokter")
            dcRujukanAsal.BoundText = .Columns("KdRujukanAsal")
            dcKualifikasiJurusan.BoundText = .Columns("KdKualifikasiJurusan")
            dcSubInstalasi.BoundText = .Columns("KdSubInstalasi")
            txtTmptRujukan.Text = .Columns("TempatRujukanAsal")
            txtAlamatRujukan.Text = .Columns("AlamatTempatRujukanAsal")
            dtpTglKirimTerima.Value = .Columns("TglTerima")
            dtpTglKembali.Value = .Columns("TglKembali")
            txtKetLainnya.Text = .Columns("Keterangan")
        End If
    End With
    Exit Sub
hell:
End Sub

Private Sub dtpTglKembali_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdsimpan.SetFocus
End Sub

Private Sub dtpTglKirimTerima_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglKembali.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click
    'Call subLoadDataGrid
End Sub

Private Sub optKirim_Click()
    Call subLoadDataGrid
    Call subKosong
End Sub

Private Sub optTerima_Click()
    Call subLoadDataGrid
    Call subKosong
End Sub

Private Sub txtAlamatRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglKirimTerima.SetFocus
End Sub

Private Sub txtKetLainnya_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtTmptRujukan.SetFocus
End Sub

Private Sub txtTmptRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtAlamatRujukan.SetFocus
End Sub

Private Function sp_RiwayatKirimTenagaPemeriksa2(f_status As String) As Boolean
    On Error GoTo hell
    sp_RiwayatKirimTenagaPemeriksa2 = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtnoriwayat.Text)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, dcDokter.BoundText)
        .Parameters.Append .CreateParameter("KdRujukanTujuan", adChar, adParamInput, 2, dcRujukanAsal.BoundText)
        .Parameters.Append .CreateParameter("KdKualifikasiJurusan", adVarChar, adParamInput, 4, dcKualifikasiJurusan.BoundText)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("NamaRujukanTujuan", adVarChar, adParamInput, 75, Trim(txtTmptRujukan.Text))
        .Parameters.Append .CreateParameter("AlamatTempatRujukan", adVarChar, adParamInput, 150, IIf(txtAlamatRujukan.Text = "", Null, Trim(txtAlamatRujukan.Text)))
        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirimTerima.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglKembali", adDate, adParamInput, , Format(dtpTglKembali.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, IIf(txtKetLainnya.Text = "", Null, Trim(txtKetLainnya.Text)))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
    

        .ActiveConnection = dbConn
        .CommandText = "Aud_RiwayatKirimTenagaPemeriksa"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_RiwayatKirimTenagaPemeriksa() As Boolean
    On Error GoTo hell
    sp_RiwayatKirimTenagaPemeriksa = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        '//Yna 2014-0808 '.Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtNoRiwayat.Text)
        
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, dcDokter.BoundText)
        .Parameters.Append .CreateParameter("KdRujukanTujuan", adChar, adParamInput, 2, dcRujukanAsal.BoundText)
        .Parameters.Append .CreateParameter("KdKualifikasiJurusan", adVarChar, adParamInput, 4, dcKualifikasiJurusan.BoundText)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("NamaRujukanTujuan", adVarChar, adParamInput, 75, Trim(txtTmptRujukan.Text))
        .Parameters.Append .CreateParameter("AlamatTempatRujukan", adVarChar, adParamInput, 150, IIf(txtAlamatRujukan.Text = "", Null, Trim(txtAlamatRujukan.Text)))
        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirimTerima.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglKembali", adDate, adParamInput, , Format(dtpTglKembali.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, IIf(txtKetLainnya.Text = "", Null, Trim(txtKetLainnya.Text)))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
    

        .ActiveConnection = dbConn
        .CommandText = "Add_RiwayatKirimTenagaPemeriksa"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_RiwayatTerimaTenagaPemeriksa2(f_status As String) As Boolean
    On Error GoTo hell
    sp_RiwayatTerimaTenagaPemeriksa2 = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtnoriwayat.Text)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, dcDokter.BoundText)
        .Parameters.Append .CreateParameter("KdRujukanTujuan", adChar, adParamInput, 2, dcRujukanAsal.BoundText)
        .Parameters.Append .CreateParameter("KdKualifikasiJurusan", adVarChar, adParamInput, 4, dcKualifikasiJurusan.BoundText)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("NamaRujukanTujuan", adVarChar, adParamInput, 75, Trim(txtTmptRujukan.Text))
        .Parameters.Append .CreateParameter("AlamatTempatRujukan", adVarChar, adParamInput, 150, IIf(txtAlamatRujukan.Text = "", Null, Trim(txtAlamatRujukan.Text)))
        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirimTerima.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglKembali", adDate, adParamInput, , Format(dtpTglKembali.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, IIf(txtKetLainnya.Text = "", Null, Trim(txtKetLainnya.Text)))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "Aud_RiwayatTerimaTenagaPemeriksa"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function


Private Function sp_RiwayatTerimaTenagaPemeriksa() As Boolean
    On Error GoTo hell
    sp_RiwayatTerimaTenagaPemeriksa = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        '//Yna 2014-0808 '.Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtNoRiwayat.Text)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, dcDokter.BoundText)
        .Parameters.Append .CreateParameter("KdRujukanTujuan", adChar, adParamInput, 2, dcRujukanAsal.BoundText)
        .Parameters.Append .CreateParameter("KdKualifikasiJurusan", adVarChar, adParamInput, 4, dcKualifikasiJurusan.BoundText)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("NamaRujukanTujuan", adVarChar, adParamInput, 75, Trim(txtTmptRujukan.Text))
        .Parameters.Append .CreateParameter("AlamatTempatRujukan", adVarChar, adParamInput, 150, IIf(txtAlamatRujukan.Text = "", Null, Trim(txtAlamatRujukan.Text)))
        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirimTerima.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglKembali", adDate, adParamInput, , Format(dtpTglKembali.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, IIf(txtKetLainnya.Text = "", Null, Trim(txtKetLainnya.Text)))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)

        .ActiveConnection = dbConn
        .CommandText = "Add_RiwayatTerimaTenagaPemeriksa"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Sub cmdCetak_Click()
    On Error GoTo errLoad

    Dim pesan As VbMsgBoxResult

    If optKirim.Value = True Then

        strSQL = "Select * From V_RiwayatKirimTenagaPemeriksa "
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then
            MsgBox "Tidak ada data ", vbInformation, "Informasi"
            Exit Sub
        End If

    ElseIf optTerima.Value = True Then
        strSQL = "Select * From V_RiwayatTerimaTenagaPemeriksa "
        Call msubRecFO(dbRst, strSQL)
        If dbRst.EOF = True Then
            MsgBox "Tidak ada data ", vbInformation, "Informasi"
            Exit Sub
        End If
    End If
    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"
    frm_cetak_LaporanRiwayatKirimTerimaTenagaPemeriksa.Show
    Exit Sub
errLoad:
End Sub


Private Function sp_Riwayat(f_status) As Boolean
    On Error GoTo hell
    sp_Riwayat = True
    Set dbcmd = New ADODB.Command
    With dbcmd

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        If txtnoriwayat = "" Then
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtnoriwayat.Text)
        End If

        .Parameters.Append .CreateParameter("TglRiwayat", adDate, adParamInput, , Format(Now, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        .Parameters.Append .CreateParameter("OutputNoRiwayat", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Riwayat"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data riwayat ", vbCritical, "Validasi"
            sp_Riwayat = False
        End If
        txtnoriwayat.Text = .Parameters("OutputNoRiwayat").Value
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function
