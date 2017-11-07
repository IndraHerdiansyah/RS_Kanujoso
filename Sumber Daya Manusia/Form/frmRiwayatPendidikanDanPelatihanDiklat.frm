VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRiwayatPendidikanDanPelatihanDiklat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Pendidikan & Pelatihan Diklat"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRiwayatPendidikanDanPelatihanDiklat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   12855
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   600
      Left            =   8160
      TabIndex        =   10
      Top             =   8040
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
      Left            =   6600
      TabIndex        =   9
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "C&etak"
      Height          =   600
      Left            =   5040
      TabIndex        =   8
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   600
      Left            =   11280
      TabIndex        =   12
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   600
      Left            =   9720
      TabIndex        =   11
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Peserta Diklat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   25
      Top             =   3240
      Width           =   12735
      Begin VB.TextBox txtIsi 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   5040
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcAsalPeserta 
         Height          =   330
         Left            =   3120
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKelompokPeserta 
         Height          =   330
         Left            =   3120
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   1695
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2990
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Diklat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   15
      Top             =   1080
      Width           =   12735
      Begin VB.TextBox txtTmptPenyelenggara 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9480
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtKetLainnya 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1680
         Width           =   10695
      End
      Begin VB.TextBox txtDeskDiklat 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   1920
         MaxLength       =   50
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   720
         Width           =   4815
      End
      Begin MSDataListLib.DataCombo dcNamaDiklat 
         Height          =   330
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
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
      Begin MSComCtl2.DTPicker dtpTglAwalDiklat 
         Height          =   330
         Left            =   1920
         TabIndex        =   2
         Top             =   1320
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
         Format          =   56033283
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpTglAkhirDiklat 
         Height          =   330
         Left            =   4560
         TabIndex        =   3
         Top             =   1320
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
         Format          =   56033283
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcPJawab 
         Height          =   330
         Left            =   9480
         TabIndex        =   5
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSDataListLib.DataCombo dcInsPenyelenggara 
         Height          =   330
         Left            =   9480
         TabIndex        =   6
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Penyelenggara"
         Height          =   210
         Index           =   7
         Left            =   7440
         TabIndex        =   23
         Top             =   1080
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instansi Penyelenggara"
         Height          =   210
         Index           =   6
         Left            =   7440
         TabIndex        =   22
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penanggung Jawab"
         Height          =   210
         Index           =   5
         Left            =   7440
         TabIndex        =   21
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan Lainnya"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1725
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         Height          =   210
         Index           =   3
         Left            =   4200
         TabIndex        =   19
         Top             =   1365
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Mulai Diklat"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deskripsi Diklat"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Diklat"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   945
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   14
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
   Begin MSDataGridLib.DataGrid dgData 
      Height          =   2535
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   12615
      _ExtentX        =   22251
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
      Picture         =   "frmRiwayatPendidikanDanPelatihanDiklat.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10920
      Picture         =   "frmRiwayatPendidikanDanPelatihanDiklat.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPendidikanDanPelatihanDiklat.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "frmRiwayatPendidikanDanPelatihanDiklat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bolTampilDetail As Boolean
Dim strNoRiwayat As String

Sub subKosong()
    On Error GoTo hell
    dcNamaDiklat.Text = ""
    txtDeskDiklat.Text = ""
    dtpTglAwalDiklat.Value = Now
    dtpTglAkhirDiklat.Value = Now
    txtKetLainnya.Text = ""
    dcPJawab.Text = ""
    dcInsPenyelenggara.Text = ""
    txtTmptPenyelenggara.Text = ""
    txtNoRiwayat.Text = ""
    bolTampilDetail = False
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subSetGrid()
    On Error GoTo hell
    With fgData
        .clear
        .Rows = 2
        .Cols = 6

        .RowHeight(0) = 400

        .TextMatrix(0, 0) = "No Urut"
        .TextMatrix(1, 0) = "1"
        .TextMatrix(0, 1) = "Nama Peserta"
        .TextMatrix(0, 2) = "Kelompok Peserta"
        .TextMatrix(0, 3) = "Asal Peserta"
        .TextMatrix(0, 4) = "KdKelompokPeserta"
        .TextMatrix(0, 5) = "KdAsalPeserta"

        .ColWidth(0) = 0
        .ColWidth(1) = 5500
        .ColWidth(2) = 3500
        .ColWidth(3) = 3100
        .ColWidth(4) = 0
        .ColWidth(5) = 0
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadDcSource()
    On Error GoTo hell
    Call msubDcSource(dcNamaDiklat, rs, "Select KdDiklat,NamaDiklat From Diklat Where StatusEnabled=1")
    Call msubDcSource(dcPJawab, rs, "Select IdPegawai,NamaLengkap From DataPegawai")
    Call msubDcSource(dcInsPenyelenggara, rs, "Select KdTempatPenyelenggara,TempatPenyelenggara From TempatPenyelenggaraDiklat Where StatusEnabled=1")
    Call msubDcSource(dcKelompokPeserta, rs, "Select KdKelompokPegawai,KelompokPegawai From KelompokPegawai Where StatusEnabled=1")
    Call msubDcSource(dcAsalPeserta, rs, "Select KdAsalPeserta,AsalPeserta From AsalPesertaDiklat Where StatusEnabled=1")
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtIsi.Left = fgData.Left

    For i = 0 To fgData.Col - 1
        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgData.Top - 7

    For i = 0 To fgData.row - 1
        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    txtIsi.Width = fgData.ColWidth(fgData.Col)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub

Private Sub subLoadDataCombo(s_DcName As Object)
    Dim i As Integer
    s_DcName.Left = fgData.Left
    For i = 0 To fgData.Col - 1
        s_DcName.Left = s_DcName.Left + fgData.ColWidth(i)
    Next i
    s_DcName.Visible = True
    s_DcName.Top = fgData.Top - 7

    For i = 0 To fgData.row - 1
        s_DcName.Top = s_DcName.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        s_DcName.Top = s_DcName.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    s_DcName.Width = fgData.ColWidth(fgData.Col)
    s_DcName.Height = fgData.RowHeight(fgData.row)

    s_DcName.Visible = True
    s_DcName.SetFocus
End Sub

Sub subLoadDataGrid()
    On Error GoTo hell
    Dim i As Integer

'    strSQL = "Select Distinct NoRiwayat,NamaDiklat,DeskripsiDiklat,TglMulai,TglSelesai,Keterangan,PenanggungJawab, " & _
'    "InstansiPenyelenggara,TempatPenyelenggara,KdDiklat,IdPegawaiPJawab,KdTempatPenyelenggara From V_RiwayatPendidikanNPelatihanDiklat"
'
    strSQL = "SELECT DISTINCT NoRiwayat, Jenis as  NamaDiklat, Kegiatan as DeskripsiDiklat, TglMulai, TglSelesai, PenanggungJawab,Institusi as InstansiPenyelenggara, " & _
             " Tempat as TempatPenyelenggara, KdDiklat, IdPegawaiPJawab " & _
             " FROM V_RiwayatPendidikanNPelatihanDiklat"
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    Set dgData.DataSource = rs
    With dgData
        For i = 1 To .Col
            .Columns(i).Width = 0
        Next i

        .Columns("NoRiwayat").Width = 1100
        .Columns("NamaDiklat").Width = 1200
        .Columns("DeskripsiDiklat").Width = 1500
        .Columns("TglMulai").Width = 1250
        .Columns("TglSelesai").Width = 1250
'        .Columns("Keterangan").Width = 1500
        .Columns("PenanggungJawab").Width = 0
        .Columns("IdPegawaiPJawab").Caption = "PenanggungJawab"
        .Columns("TempatPenyelenggara").Width = 1300
        .Columns("InstansiPenyelenggara").Width = 1000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub CmdBatal_Click()
    Call subKosong
    Call subSetGrid
    Call subLoadDcSource
    Call subLoadDataGrid
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo errLoad

    Dim pesan As VbMsgBoxResult

    strSQL = "Select Distinct NoRiwayat,NamaDiklat,DeskripsiDiklat,TglMulai,TglSelesai,Keterangan,PenanggungJawab, " & _
    "InstansiPenyelenggara,TempatPenyelenggara,KdDiklat,IdPegawaiPJawab,KdTempatPenyelenggara From V_RiwayatPendidikanNPelatihanDiklat"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        MsgBox "Tidak ada data ", vbInformation, "Informasi"
        Exit Sub
    End If

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"
    frm_cetak_LaporanRiwayatDiklat.Show
    Exit Sub
errLoad:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoRiwayat.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatDiklatPelatihan WHERE NoRiwayat='" & txtNoRiwayat.Text & "' "
    Call msubRecFO(rs, strSQL)
    strSQL = "DELETE FROM Riwayat WHERE NoRiwayat='" & txtNoRiwayat.Text & "' "
    Call msubRecFO(rs, strSQL)

    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call CmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If dcNamaDiklat.Text <> "" Then
        If Periksa("datacombo", dcNamaDiklat, "Nama Diklat Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcPJawab.Text = "" Then
        If Periksa("datacombo", dcPJawab, "Penanggung Jawab Tidak Terdaftar") = False Then Exit Sub
'         MsgBox "datacombo", dcPJawab, "Penanggung Jawab Tidak Terdaftar", vbInformation
         dcPJawab.SetFocus
         Exit Sub
    End If
    If dcInsPenyelenggara.Text <> "" Then
        If Periksa("datacombo", dcInsPenyelenggara, "Instansi Penyelenggara Tidak Terdaftar") = False Then Exit Sub
    End If
    
    
    If Periksa("datacombo", dcNamaDiklat, "Silahkan pilih Nama Diklat ") = False Then Exit Sub
'    If Periksa("datacombo", dcPJawab, "Silahkan pilih Penanggung jawab ") = False Then Exit Sub
    If Periksa("datacombo", dcInsPenyelenggara, "Silahkan pilih Instalasi penyelenggara ") = False Then Exit Sub
    If fgData.row = 1 Then
        MsgBox "Data peserta harus diisi", vbCritical
'        fgData.row = 1: fgData.SetFocus
        Exit Sub
    Else
        If sp_RiwayatDiklatPelatihan = False Then Exit Sub
    End If

    With fgData
        For i = 1 To fgData.Rows
            If .TextMatrix(i, 1) = "" Then GoTo keluar_
            If .TextMatrix(i, 2) = "" Then GoTo keluar_
            If .TextMatrix(i, 3) = "" Then GoTo keluar_
            If sp_RiwayatPesertaDiklat(i, .TextMatrix(i, 1), .TextMatrix(i, 4), .TextMatrix(i, 5), "A") = False Then Exit Sub
        Next i
    End With
keluar_:
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Call CmdBatal_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAsalPeserta_Change()
    On Error GoTo errLoad
    fgData.TextMatrix(fgData.row, 3) = dcAsalPeserta.Text
    fgData.TextMatrix(fgData.row, 5) = dcAsalPeserta.BoundText
    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub dcAsalPeserta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcAsalPeserta_Change
        dcAsalPeserta.Visible = False

        With fgData
            If .TextMatrix(.Rows - 1, 2) = "" Then
                .row = .Rows - 1
                .Col = 0
            Else

                .SetFocus
                .Rows = .Rows + 1
                .row = .row + 1
                .Col = 0

            End If
        End With

        fgData.Col = 1
        fgData.SetFocus
    End If
End Sub

Private Sub dcAsalPeserta_LostFocus()
    dcAsalPeserta.Visible = False
End Sub

Private Sub dcInsPenyelenggara_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcInsPenyelenggara.MatchedWithList = True Then txtTmptPenyelenggara.SetFocus
        strSQL = "Select KdTempatPenyelenggara,TempatPenyelenggara From TempatPenyelenggaraDiklat WHERE (TempatPenyelenggara LIKE '%" & dcInsPenyelenggara.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcInsPenyelenggara.BoundText = rs(0).Value
        dcInsPenyelenggara.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKelompokPeserta_Change()
    On Error GoTo errLoad
    fgData.TextMatrix(fgData.row, 2) = dcKelompokPeserta.Text
    fgData.TextMatrix(fgData.row, 4) = dcKelompokPeserta.BoundText
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKelompokPeserta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcKelompokPeserta_Change
        dcKelompokPeserta.Visible = False
        fgData.Col = 3
        fgData.SetFocus
    End If

'On Error GoTo hell
'    If KeyAscii = 39 Then KeyAscii = 0
'    If KeyAscii = 13 Then
'        If dcKelompokPeserta.MatchedWithList = True Then fgData.SetFocus
'        strSQL = "Select KdKelompokPegawai,KelompokPegawai From KelompokPegawai Where (KelompokPegawai LIKE '%" & dcKelompokPeserta.Text & "%') And StatusEnabled=1"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = True Then Exit Sub
'        dcKelompokPeserta.BoundText = rs(0).Value
'        dcKelompokPeserta.Text = rs(1).Value
'    End If
'    Exit Sub
'hell:
'    Call msubPesanError
End Sub

Private Sub dcKelompokPeserta_LostFocus()
    dcKelompokPeserta.Visible = False
End Sub

Private Sub dcNamaDiklat_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcNamaDiklat.MatchedWithList = True Then txtDeskDiklat.SetFocus
        strSQL = "Select KdDiklat,NamaDiklat From Diklat WHERE (NamaDiklat LIKE '%" & dcNamaDiklat.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcNamaDiklat.BoundText = rs(0).Value
        dcNamaDiklat.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcPJawab_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcPJawab.MatchedWithList = True Then dcInsPenyelenggara.SetFocus
        strSQL = "Select IdPegawai,NamaLengkap From DataPegawai WHERE (NamaLengkap LIKE '%" & dcPJawab.Text & "%') "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPJawab.BoundText = rs(0).Value
        dcPJawab.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcPJawab_LostFocus()
On Error Resume Next
    
    Set rs = Nothing
    strSQL = "Select IdPegawai,NamaLengkap From DataPegawai where NamaLengkap like '%" & dcPJawab.Text & "%'"
    Call msubRecFO(rs, strSQL)
    
    If rs.EOF = True Then
        MsgBox "Penanggung Jawab Tidak Terdaftar", vbInformation
        dcPJawab.Text = ""
        dcPJawab.SetFocus
    Exit Sub
    End If

Exit Sub
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo hell
    With dgData
        Call subSetGrid
        If .ApproxCount <= 0 Then Exit Sub
        bolTampilDetail = True
        txtNoRiwayat.Text = .Columns("NoRiwayat")
        dcNamaDiklat.BoundText = .Columns("KdDiklat")
        txtDeskDiklat.Text = .Columns("DeskripsiDiklat")
        dtpTglAwalDiklat.Value = .Columns("TglMulai")
        dtpTglAkhirDiklat.Value = .Columns("TglSelesai")
        txtKetLainnya.Text = .Columns("Keterangan")
        If IsNull(.Columns("TempatPenyelenggara")) Then txtTmptPenyelenggara.Text = "" Else txtTmptPenyelenggara.Text = .Columns("TempatPenyelenggara")
        dcInsPenyelenggara.BoundText = .Columns("KdTempatPenyelenggara")
        dcPJawab.Text = .Columns("PenanggungJawab").Value '//yayang.agus 2014-08-13
        'dcPJawab.Text = .Columns(10).Value
    End With
    Exit Sub
hell:
End Sub

Private Sub dtpTglAkhirDiklat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKetLainnya.SetFocus
End Sub

Private Sub dtpTglAwalDiklat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglAkhirDiklat.SetFocus
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    txtIsi.Text = ""

    Select Case fgData.Col

        Case 1 'Nama peserta
            txtIsi.MaxLength = 0
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)

        Case 2 'kelompok peserta
            Call subLoadDataCombo(dcKelompokPeserta)

        Case 3 'asal peserta
            Call subLoadDataCombo(dcAsalPeserta)

    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call CmdBatal_Click

End Sub

Private Sub txtDeskDiklat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglAwalDiklat.SetFocus
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case fgData.Col
            Case 0

            Case 1
                fgData.TextMatrix(fgData.row, 1) = txtIsi.Text
                fgData.SetFocus
                fgData.Col = 2

            Case 3

        End Select

        txtIsi.Visible = False

        If fgData.RowPos(fgData.row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If

    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)

    Select Case KeyCode
        Case 13
            If fgData.TextMatrix(fgData.row, 2) = "" Then Exit Sub
            Call subLoadText
            txtIsi.Text = Trim(fgData.TextMatrix(fgData.row, fgData.Col))
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)

        Case vbKeyDelete

            If fgData.row = fgData.Rows - 1 Then Exit Sub
            fgData.RemoveItem fgData.row

    End Select
End Sub

Private Sub txtKetLainnya_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPJawab.SetFocus
End Sub

Private Sub txtNoRiwayat_Change()
    On Error GoTo hell
    Dim i As Integer
    If bolTampilDetail = False Then Exit Sub

    strSQL = "Select * From V_RiwayatPendidikanNPelatihanDiklat Where NoRiwayat='" & txtNoRiwayat.Text & "'"
    Set dbRst = Nothing
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then Exit Sub

    With fgData
        For i = 1 To dbRst.RecordCount
            .TextMatrix(i, 0) = IIf(IsNull(dbRst("NoUrut")), "", dbRst("NoUrut"))
            .TextMatrix(i, 1) = IIf(IsNull(dbRst("NamaPeserta")), "", dbRst("NamaPeserta"))
            .TextMatrix(i, 2) = IIf(IsNull(dbRst("KelompokPegawai")), "", dbRst("KelompokPegawai"))
            .TextMatrix(i, 3) = IIf(IsNull(dbRst("AsalPeserta")), "", dbRst("AsalPeserta"))
            .TextMatrix(i, 4) = IIf(IsNull(dbRst("KdKelompokPegawai")), "", dbRst("KdKelompokPegawai"))
            .TextMatrix(i, 5) = IIf(IsNull(dbRst("KdAsalPeserta")), "", dbRst("KdAsalPeserta"))

            dbRst.MoveNext
            .Rows = .Rows + 1
        Next i
        .row = 1
    End With
    bolTampilDetail = False
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtTmptPenyelenggara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then fgData.SetFocus
End Sub

Private Function sp_RiwayatDiklatPelatihan() As Boolean
    On Error GoTo hell
    sp_RiwayatDiklatPelatihan = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDiklat", adVarChar, adParamInput, 5, dcNamaDiklat.BoundText)
        .Parameters.Append .CreateParameter("DeskripsiDiklat", adVarChar, adParamInput, 150, IIf(txtDeskDiklat.Text = "", Null, Trim(txtDeskDiklat.Text)))
        .Parameters.Append .CreateParameter("TglMulai", adDate, adParamInput, , Format(dtpTglAwalDiklat.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglSelesai", adDate, adParamInput, , Format(dtpTglAkhirDiklat.Value, "yyyy/MM/dd HH:mm:ss"))
        '.Parameters.Append .CreateParameter("PegawaiPJawab", adVarChar, adParamInput, 150, dcPJawab.Text)
        .Parameters.Append .CreateParameter("PegawaiPJawab", adVarChar, adParamInput, 150, dcPJawab.BoundText) '//Yna 2014-0808
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, IIf(txtKetLainnya.Text = "", Null, Trim(txtKetLainnya.Text)))
        .Parameters.Append .CreateParameter("KdTempatPenyelenggara", adTinyInt, adParamInput, , dcInsPenyelenggara.BoundText)
        .Parameters.Append .CreateParameter("TempatPenyelenggara", adVarChar, adParamInput, 30, IIf(txtTmptPenyelenggara.Text = "", Null, Trim(txtTmptPenyelenggara.Text)))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("output", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "Add_RiwayatDiklatPelatihan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
            sp_RiwayatDiklatPelatihan = False
        Else
            strNoRiwayat = .Parameters("output").Value
        End If

        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_RiwayatPesertaDiklat(ByVal iNoUrut As Integer, strNamaPegawai As String, strKdKelompok As String, strKdAsalPeserta As String, strStatus As String) As Boolean
    On Error GoTo hell
    sp_RiwayatPesertaDiklat = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, strNoRiwayat)
        .Parameters.Append .CreateParameter("NoUrut", adSmallInt, adParamInput, , iNoUrut)
        .Parameters.Append .CreateParameter("NamaPegawai", adVarChar, adParamInput, 50, strNamaPegawai)
        .Parameters.Append .CreateParameter("KdKelompokPegawai", adChar, adParamInput, 2, strKdKelompok)
        .Parameters.Append .CreateParameter("KdAsalPeserta", adChar, adParamInput, 2, strKdAsalPeserta)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatus)

        .ActiveConnection = dbConn
        .CommandText = "AUD_PesertaDiklatPelatihan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
            sp_RiwayatPesertaDiklat = False
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

