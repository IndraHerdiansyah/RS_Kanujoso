VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Medifirst2000 - Sumber Daya Manusia (Human Resource)"
   ClientHeight    =   8100
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14565
   Icon            =   "MDIUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIUtama.frx":0CCA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   8040
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   7560
      Top             =   4440
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7845
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6959
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6959
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "24/03/2017"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:41"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
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
   Begin VB.Menu mnuBerkas 
      Caption         =   "&Berkas"
      Begin VB.Menu mnudata 
         Caption         =   "Data"
         Begin VB.Menu mnuPersonalPegawai 
            Caption         =   "Personal Pegawai"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuline3 
            Caption         =   "-"
         End
         Begin VB.Menu mnAbsensiPegawai 
            Caption         =   "Absensi "
            Begin VB.Menu mnuabsen 
               Caption         =   "Absensi Pegawai"
            End
            Begin VB.Menu mnuPIN 
               Caption         =   "Nomor PIN"
            End
            Begin VB.Menu mnuCekAbsensi 
               Caption         =   "Monitoring Absensi Pegawai"
            End
         End
         Begin VB.Menu mnumasterabsensi 
            Caption         =   "Master Absensi"
            Begin VB.Menu mnudataHariLibur 
               Caption         =   "Hari Libur"
            End
            Begin VB.Menu mnuShift2 
               Caption         =   "Shift2"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuShift 
               Caption         =   "Shift"
            End
            Begin VB.Menu batasabsenjadwal 
               Caption         =   "-"
            End
            Begin VB.Menu mnuJadwalKerja 
               Caption         =   "Jadwal Kerja"
            End
            Begin VB.Menu batasstatusabsen 
               Caption         =   "-"
            End
            Begin VB.Menu mnustatusabsensi 
               Caption         =   "Status Absen"
            End
         End
         Begin VB.Menu mnbatas 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataMasterPegawai 
            Caption         =   "Master Pegawai"
            Begin VB.Menu mnUJenPeg 
               Caption         =   "Jenis Pegawai"
            End
            Begin VB.Menu mnuKatpeg 
               Caption         =   "Kategory Pegawai"
            End
            Begin VB.Menu mnuKelPeg 
               Caption         =   "Kelompok Pegawai"
            End
            Begin VB.Menu mnuTypePeg 
               Caption         =   "Type Pegawai"
            End
            Begin VB.Menu mnubasat1 
               Caption         =   "-"
            End
            Begin VB.Menu mnuJabatanMaster 
               Caption         =   "Jenis && Nama Jabatan"
            End
            Begin VB.Menu mnuJenjangJabatan 
               Caption         =   "Jenjang Jabatan"
            End
            Begin VB.Menu mnuPangGol 
               Caption         =   "Pangkat && Golongan"
            End
            Begin VB.Menu mnuKuali 
               Caption         =   "Pendidikan && Kualifikasi Jurusan"
            End
            Begin VB.Menu mnuTitle 
               Caption         =   "Gelar"
            End
         End
         Begin VB.Menu garisF 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMasterF 
            Caption         =   "Data Master Jabatan Fungsional Pegawai"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMasterFu 
            Caption         =   "Data Master Pelayanan Fungsional"
            Visible         =   0   'False
         End
         Begin VB.Menu ww 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMstrGAP 
            Caption         =   "Master GAP Kopetensi"
         End
         Begin VB.Menu mnuKinerjaPegawai 
            Caption         =   "Master Kinerja"
            Begin VB.Menu MS 
               Caption         =   "Master Kinerja"
            End
            Begin VB.Menu mnuSasaranKinerja 
               Caption         =   "Sasaran Kinerja"
            End
            Begin VB.Menu mnuNilaiKinerja 
               Caption         =   "Penilaian Kinerja"
            End
            Begin VB.Menu mnuRekapKinerja 
               Caption         =   "Rekap Kinerja"
            End
         End
         Begin VB.Menu mnuline2 
            Caption         =   "-"
         End
         Begin VB.Menu mnstatuspegawai 
            Caption         =   "Master Status"
         End
         Begin VB.Menu mnuline 
            Caption         =   "-"
         End
         Begin VB.Menu mnudatapenunjang 
            Caption         =   "Master Penunjang"
         End
         Begin VB.Menu mnuBatas1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPenilaianPegawai 
            Caption         =   "Penilaian Pegawai"
         End
         Begin VB.Menu mnuDaftarPenilaianPegawai2 
            Caption         =   "Daftar Penilaian Pegawai"
            Visible         =   0   'False
         End
         Begin VB.Menu KomponenGaji 
            Caption         =   "Master Komponen Gaji"
         End
         Begin VB.Menu PembayaranGajiPegawai 
            Caption         =   "Perhitungan Gaji Pegawai"
         End
         Begin VB.Menu bataskomponen 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMasterDataInsentif 
            Caption         =   "Master Insentif"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHitungInsentif 
            Caption         =   "Perhitungan Insentif Pegawai"
            Visible         =   0   'False
         End
         Begin VB.Menu batasgarisnilai 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDataKomponenIndex 
            Caption         =   "Komponen Index"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDataPerhitunganIndex 
            Caption         =   "Perhitungan Index"
            Visible         =   0   'False
         End
         Begin VB.Menu batasusulan 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTransaksiRiwayatUsulan 
            Caption         =   "Riwayat Usulan Pegawai"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRiwayatKirimdanTerimaTenagaMedis 
            Caption         =   "Riwayat Kirim dan Terima Tenaga Medis"
         End
         Begin VB.Menu RiwayatPendidikandanLatihan 
            Caption         =   "Riwayat Pendidikan dan Latihan"
         End
      End
      Begin VB.Menu mnuBatas2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingPrinter 
         Caption         =   "Setting Printer"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuGantiKataKunci 
         Caption         =   "Ganti Kata Kunci"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuBatas3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuSelesai 
         Caption         =   "Selesai"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuInformasi 
      Caption         =   "&Informasi"
      Begin VB.Menu mnuNotifikasi 
         Caption         =   "Notifikasi"
      End
      Begin VB.Menu mnuGAP 
         Caption         =   "Daftar Mapping GAP Kopetensi"
      End
      Begin VB.Menu mnuInformasiDaftarPegawai 
         Caption         =   "Informasi Daftar Pegawai"
      End
      Begin VB.Menu mnuDUK 
         Caption         =   "Daftar Urut Kepangkatan (DUK)"
      End
      Begin VB.Menu mnuDaftarPenilaianPegawai 
         Caption         =   "Daftar Penilaian Pegawai"
      End
      Begin VB.Menu mnuDaftarUsulan 
         Caption         =   "Daftar Riwayat Usulan Pegawai"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuJasaPelayananPegawai 
         Caption         =   "Daftar Jasa Pelayanan Pegawai"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudaftarTaspen 
         Caption         =   "Daftar Usulan Penerbitan TASPEN"
         Visible         =   0   'False
      End
      Begin VB.Menu garisPelayananPegawai 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDaftarLayananPegawai 
         Caption         =   "Daftar Pelayanan Pegawai"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuriwayatpangkat 
         Caption         =   "Kenaikan Pangkat"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuriwayatgaji 
         Caption         =   "Kenaikan Gaji Berkala"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBts 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPayroll 
         Caption         =   "Daftar Penghasilan Pegawai Non PNS"
      End
      Begin VB.Menu mnMinitoringJadwalKerjaRuangan 
         Caption         =   "Monitoring Jadwal Kerja Ruangan"
      End
   End
   Begin VB.Menu mInventory 
      Caption         =   "In&ventory"
      Begin VB.Menu mPemesananBarang 
         Caption         =   "Pemesanan Barang"
      End
      Begin VB.Menu mnuMonitoringStokBarangNM 
         Caption         =   "Monitoring Stok Barang Non Medis"
      End
      Begin VB.Menu mnuPemakaianBahandanAlat 
         Caption         =   "Pemakaian Bahan dan Alat"
         Visible         =   0   'False
      End
      Begin VB.Menu grs1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mBarangMedis 
         Caption         =   "Barang Medis"
         Visible         =   0   'False
         Begin VB.Menu mStokBarang 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu mClosingStok 
            Caption         =   "Closing Stok"
            Begin VB.Menu mCetakLembarInput 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu mInputStokOpn 
               Caption         =   "Input Stok Opname"
            End
            Begin VB.Menu mNilaiPersediaan 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu grs2 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiPemesananBarang 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu POAKaryawan 
            Caption         =   "Informasi Pemakaian Barang"
         End
         Begin VB.Menu mLapSaldoBarang 
            Caption         =   "Laporan Saldo Barang"
         End
      End
      Begin VB.Menu mBarangNM 
         Caption         =   "Barang Non Medis"
         Begin VB.Menu mStokBarangNM 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu mKondisiBarangNM 
            Caption         =   "Kondisi Barang"
         End
         Begin VB.Menu mMutasiBarangNM 
            Caption         =   "Mutasi Barang"
         End
         Begin VB.Menu mRekapitulasiTranBrgNM 
            Caption         =   "Rekapitulasi Transaksi Barang"
            Visible         =   0   'False
         End
         Begin VB.Menu mClosingStokNM 
            Caption         =   "Closing Stok"
            Begin VB.Menu mCetakLembarInputNM 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu mInputStokOpnameNM 
               Caption         =   "Input Stok Opname"
            End
            Begin VB.Menu mNilaiPersediaanNM 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu ln1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mInfoPemesananBrgNM 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu mLapSaldoBarangNM 
            Caption         =   "Laporan Saldo Barang"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuLaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnucetakabsen 
         Caption         =   "Absensi Pegawai"
      End
      Begin VB.Menu m_slipinsentifpegawai 
         Caption         =   "Slip Insentif Pegawai"
      End
      Begin VB.Menu gLapAbsen 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLaporanDaftarPegawai 
         Caption         =   "Daftar Pegawai"
      End
      Begin VB.Menu mnuLaporanIndexPegawai 
         Caption         =   "Index Pegawai"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLaporanBukuCVPegawai 
         Caption         =   "CV Pegawai"
      End
      Begin VB.Menu mnureqPNS 
         Caption         =   "Rekap Pegawai Berdasarkan Pendidikan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBezet 
         Caption         =   "Rekap Pegawai Berdasarkan Jabatan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBezettingPangkat 
         Caption         =   "Rekap Pegawai Berdasarkan Pangkat"
         Visible         =   0   'False
      End
      Begin VB.Menu batasriwayat 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuriwayatperjalanandinas 
         Caption         =   "Riwayat Perjalanan Dinas Pegawai"
         Visible         =   0   'False
      End
      Begin VB.Menu batasmnuLaporanbaru 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu RekapLaporanGaji 
         Caption         =   "Rekap Laporan Gaji"
         Visible         =   0   'False
      End
      Begin VB.Menu batasgaji 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mn 
         Caption         =   "Laporan Daftar Pegawai"
         Begin VB.Menu mnKlasifikasi 
            Caption         =   "Laporan Daftar Pegawai Menurut Klasifikasi"
         End
         Begin VB.Menu mnKlasifikasiDetail 
            Caption         =   "Laporan Daftar Pegawai Menurut Klasifikasi Detail"
         End
         Begin VB.Menu mnJeniskelamin 
            Caption         =   "Laporan Daftar Pegawai Menurut Jenis Kelamin"
         End
         Begin VB.Menu mnLapGol 
            Caption         =   "Laporan Daftar Pegawai Menurut Golongan"
         End
         Begin VB.Menu mnLapPegBerPen 
            Caption         =   "Laporan Daftar Pegawai Menurut Pendidikan"
         End
      End
      Begin VB.Menu mnuLapJmlPegawai 
         Caption         =   "Laporan Bulanan Jumlah PNS, CPNS && TPHL (FORM I.1)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLaporanBulananB 
         Caption         =   "Laporan Bulanan Jumlah CPNS && PNS berdasarkan Pendidikan (FORM I.2)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnubyGaji 
         Caption         =   "Laporan Bulanan Jumlah PNS dan CPNS berdasarkan Gaji (FORM I.3)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapBulananPNS 
         Caption         =   "Laporan Bulanan Jumlah PNS berdasarkan Jabatan (FORM I.4)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapNominatifPegawai 
         Caption         =   "Laporan Tahunan Daftar Nominatif Pegawai (FORM I.5)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapUsulATPHL 
         Caption         =   "Laporan Tahunan Usulan Pengangkatan TPHL (FORM I.6)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapUsulanTPHLBerhenti 
         Caption         =   "Laporan Usulan Pemberhentian TPHL (FORM I.7)"
         Visible         =   0   'False
      End
      Begin VB.Menu batasFORM 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnubukunominatif 
         Caption         =   "Buku Nominatif TPHL (FORM.II.1)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapUsulanPengangkatanPNS 
         Caption         =   "Buku Catatan Tentang Pengangkatan PNS (II.2)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapKenaikanGajiBerkala 
         Caption         =   "Buku Penjagaan Kenaikan Gaji Berkala (FORM II.3)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapKenaikanPangkat 
         Caption         =   "Buku Penjagaan Kenaikan Pangkat (FORM II.4)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLaporanMutasiPegawai 
         Caption         =   "Buku Catatan Mutasi/Pindah PNS (FORMII.5)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLaporanRiwayatTugasBelajar 
         Caption         =   "Buku Catatan PNS Tugas (FORM II.6)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapHukuman 
         Caption         =   "Buku Catatan Hukuman PNS (FORM II.7)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapPenonaktifanPegawai 
         Caption         =   "Buku Catatan Non Aktif PNS (FORM II.8)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapUsulanPensiun 
         Caption         =   "Buku Catatan Usulan Pensiun PNS (FORM II.9)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBukuRegisterPensiun 
         Caption         =   "Buku Register Pensiun (FORM II.10)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuriwayatstatus 
         Caption         =   "Buku Catatan Cuti PNS (FORM II.11)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLapUsulanTaperum 
         Caption         =   "Buku Catatan TAPERUM PNS (FORM II.12)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLaporanRiwayatPrestasi 
         Caption         =   "Buku Catatan Prestasi Pegawai (FORM II.13)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuBantuan 
      Caption         =   "Bantuan"
      Begin VB.Menu mnJobOrder 
         Caption         =   "Job Order"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTentangMedifirst2000 
         Caption         =   "Tentang Medifirst2000"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuJadwalTetap 
         Caption         =   "Jadwal Tetap"
         Begin VB.Menu mnuPagi 
            Caption         =   "Pagi"
         End
         Begin VB.Menu mnuSiang 
            Caption         =   "Siang"
         End
         Begin VB.Menu mnuMalam 
            Caption         =   "Malam"
         End
         Begin VB.Menu mnuReguler 
            Caption         =   "Reguler"
         End
         Begin VB.Menu mnuGrs 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHapus 
            Caption         =   "Hapus"
         End
      End
   End
End
Attribute VB_Name = "MDIUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sepuh As Boolean
Dim TimerNotif As Double
Dim LamaNotif As Double

Private Sub subIsiJadwalTetap(Optional ByVal KodeShift As String, Optional ByVal Hapus As Boolean)
    Dim i As Integer
    Dim intRowPilih As Integer, intColPilih As Integer

    With frmJadwalKerja.fgJadwalKerja
        intRowPilih = .row
        intColPilih = .Col
        For i = 1 To .Cols - 1
            .row = 1
            .Col = i
            If Hapus Then
                .TextMatrix(intRowPilih, i) = ""
                GoTo jump
            End If
            If .CellBackColor = vbRed Then
                .TextMatrix(intRowPilih, i) = "L"
            Else
                .TextMatrix(intRowPilih, i) = KodeShift
            End If
jump:
        Next
        .row = intRowPilih
        .Col = intColPilih
    End With
End Sub

Private Sub KomponenGaji_Click()
    frmKomponenGaji.Show
End Sub

Private Sub m_slipinsentifpegawai_Click()
    frmSlipInsentifPegawai.Show
End Sub

Private Sub mCetakLembarInput_Click()
    mstrKdKelompokBarang = "02"
    frmDaftarCetakInputStokOpname.Show
End Sub

Private Sub mCetakLembarInputNM_Click()
    mstrKdKelompokBarang = "01"
    frmDaftarCetakInputStokOpnameNM.Show
End Sub

Private Sub MDIForm_Load()
    strSQL = "SELECT * FROM DataPegawai WHERE IdPegawai = '" & strIDPegawaiAktif & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    strNmPegawai = rs.Fields("NamaLengkap").Value
    Set rs = Nothing
    mnuLogOff.Caption = "Log Off..." & strNmPegawai
    StatusBar1.Panels(1).Text = "Nama User : " & strNmPegawai & "                 Database : " & strDatabaseName
    StatusBar1.Panels(2).Text = "Nama Ruangan : " & mstrNamaRuangan
    StatusBar1.Panels(5).Text = "Nama Komputer : " & strNamaHostLocal
    
    
    strSQL = "SELECT Prefix, Value, Keterangan From SettingGlobal WHERE (Prefix = 'PathFileSDM')"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount = 0 Then
        MsgBox "Path file upload SDM belum di setting", vbInformation, "..:."
        Exit Sub
    Else
        mstrPathFileSDM = Replace(rs!Value, "/", "\")
    End If
    
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim q As String
    If sepuh = True Then
        q = MsgBox("Ganti Pemakai ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then
            Unload frmLogin
            Cancel = 1
        Else
            Cancel = 0
            frmLogin.Show
        End If
        sepuh = False
    Else
        q = MsgBox("Tutup Aplikasi ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then

            Unload frmLogin
            Cancel = 1
        Else
            Cancel = 0
        End If
    End If
End Sub

Private Sub mInfoPemesananBrgNM_Click()
    mstrKdKelompokBarang = "01"
    frmInfoPesanBarangNM.Show
End Sub

Private Sub MInformasiPemesananBarang_Click()
    mstrKdKelompokBarang = "02"
    frmInfoPesanBarang.Show
End Sub

Private Sub mInputStokOpn_Click()
    mstrKdKelompokBarang = "02"
    frmStokOpname.Show
End Sub

Private Sub mInputStokOpnameNM_Click()
    mstrKdKelompokBarang = "01"
    frmStokOpnameNM.Show
End Sub

Private Sub mKondisiBarangNM_Click()
    frmKondisiBarangNM.Show
End Sub

Private Sub mLaporanAbsensi_Click()
    frmAbsensiPegawai.Show
End Sub

Private Sub mLapSaldoBarang_Click()
    frmLaporanSaldoBarangMedis_v3.Show
End Sub

Private Sub mLapSaldoBarangNM_Click()
    frmLaporanSaldoBarangNM_v3.Show
End Sub

Private Sub mMutasiBarangNM_Click()
    frmMutasiBarangNM.Show
End Sub

Private Sub mNilaiPersediaan_Click()
    mstrKdKelompokBarang = "02"
    frmNilaiPersediaan.Show
End Sub

Private Sub mNilaiPersediaanNM_Click()
    mstrKdKelompokBarang = "01"
    frmNilaiPersediaanNM.Show
End Sub

Private Sub mnsetting_Click()
    frmSetting.Show
End Sub

Private Sub mnJeniskelamin_Click()
frmLaporanDaftarPegawaiMenurutJk.Show
End Sub

Private Sub mnJobOrder_Click()
    frmJobList.Show
End Sub

Private Sub mnKlasifikasi_Click()
frmLaporanDaftarPegawaiMenurutKlasifikasi.Show
End Sub

Private Sub mnKlasifikasiDetail_Click()
frmLaporanDaftarPegawaiMenurutKlasifikasiDetail.Show
End Sub

Private Sub mnLapGol_Click()
frmLaporanDaftarPegawaiMenurutGolongan.Show
End Sub

Private Sub mnLapPegBerPen_Click()
frmLaporanDaftarPegawaiMenurutPendidikan.Show
End Sub

Private Sub mnMinitoringJadwalKerjaRuangan_Click()
    frmMonitoringJadwalKerjaRuangan.Show
End Sub

Private Sub mnstatuspegawai_Click()
    frmStatusPegawaiNew.Show
End Sub

Private Sub mnTempatBertugas_Click()
    frmRiwayatTempatBertugas.Show
End Sub

Private Sub mnuabsen_Click()
    frmAbsensiPegawai_OffLine.Show
End Sub

Private Sub mnuBezet_Click()
    On Error GoTo hell
    Dim pesan As VbMsgBoxResult
    strSQL = "SELECT DISTINCT * " & _
    " FROM V_RekapPegawaiBerdasarkanJabatan "

    Call msubRecFO(rs, strSQL)

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    frmCetakDaftarPegawaibyJabatan.Show
    Exit Sub
hell:
End Sub

Private Sub mnuBezettingPangkat_Click()
    On Error GoTo hell
    Dim pesan As VbMsgBoxResult
    strSQL = "SELECT DISTINCT * " & _
    " FROM V_RekapPegawaiBerdasarkanPangkat "

    Call msubRecFO(rs, strSQL)

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    frmCetakDaftarPegawaibyPangkat.Show
    Exit Sub
hell:
End Sub

Private Sub mnubukunominatif_Click()
    frmLaporanRiwayatNOMINATIFTPHL.Show
End Sub

Private Sub mnuBukuRegisterPensiun_Click()
    frmLaporanRegisterPensiun.Show
End Sub

Private Sub mnubyGaji_Click()
    frmLaporanBulananPegawaiD.Show
End Sub

Private Sub mnuCekAbsensi_Click()
    frmCekAbsensi.Show
End Sub

Private Sub mnucetakabsen_Click()
    frmLaporanDetailAbsensi.Show
End Sub

Private Sub mnuDaftarLayananPegawai_Click()
    frmDaftarPelayananPegawai.Show
End Sub

Private Sub mnuDaftarPenilaianPegawai_Click()
    frmDaftarPenilaianPegawai.Show
End Sub

Private Sub mnuDaftarPenilaianPegawai2_Click()
    frmDaftarPenilaianPegawai.Show
End Sub

Private Sub mnudaftarTaspen_Click()
    On Error GoTo hell
    Dim pesan As VbMsgBoxResult
    strSQL = "Select * from V_SuratKeteranganCPNSKePNS order by NamaLengkap"
    Call msubRecFO(rs, strSQL)
    If rs.BOF = True Then MsgBox "Tidak ada data", vbInformation, "": Exit Sub
    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"
    frm_cetak_SuratKeteranganCPNSKePNS.Show
hell:
End Sub

Private Sub mnuDaftarUsulan_Click()
    frmDaftarUsulanPegawaiMassal.Show
End Sub

Private Sub mnudataHariLibur_Click()
    frmDataTanggal.Show
End Sub

Private Sub mnuDataKomponenIndex_Click()
    frmDataKomponenIndex2.Show
End Sub

Private Sub mnuDataPenunjang_Click()
    frmMasterDataPenunjang.Show
End Sub

Private Sub mnuDataPerhitunganIndex_Click()
    frmDataPerhitunganIndex.Show
End Sub

Private Sub mnuDataPersonalPegawai_Click()
    frmDataPegawai.Show
End Sub

Private Sub mnuDetailPegawai_Click()
    frmDetailPegawai.Show
End Sub

Private Sub mnuGajiPegawai_Click()
    frmRiwayatGaji.Show
End Sub

Private Sub mnuGajiPokok_Click()
    frmRiwayatGaji.Show
End Sub

Private Sub mnufinger_Click()
    frmKFRS.Show
End Sub

Private Sub mnudatatgl_Click()
    frmDataTanggal.Show
End Sub

Private Sub mnuDUK_Click()
    frmDUK.Show
End Sub

Private Sub mnuGantiKataKunci_Click()
    frmGantiPassword.Show
End Sub

Private Sub mnuGAP_Click()
frmDaftarSchedule.Show
End Sub

Private Sub mnuHapus_Click()
    Call subIsiJadwalTetap(, True)
End Sub

Private Sub mnuHitungInsentif_Click()
    frmDataPerhitunganInsentif.Show
End Sub

Private Sub mnuInformasiDaftarPegawai_Click()
    frmInformasiDataPegawai.Show
End Sub

Private Sub mnuKeluargaPegawai_Click()
    frmRiwayatKeluargaPegawai.Show
End Sub

Private Sub mnuJabatanMaster_Click()
    frmJabatanPegawai.Show
End Sub

Private Sub mnujadwalkerja_Click()
    frmJadwalKerja.Show
End Sub

Private Sub mnuJasaPelayananPegawai_Click()
    frmLaporanJasaPelayanan.Show
End Sub

Private Sub mnuJenisNilai_Click()
    frmJenisNilai.Show
End Sub

Private Sub mnuJenjangJabatan_Click()
    frmJenjangJabatan.Show
End Sub

Private Sub mnUJenPeg_Click()
    frmJenisPegawai.Show
End Sub

Private Sub mnuKatpeg_Click()
    frmKategoryPegawai.Show
End Sub

Private Sub mnuKelPeg_Click()
    frmKelompokPegawai.Show
End Sub

Private Sub mnuKuali_Click()
    frmKualifikasiJurusan.Show
End Sub

Private Sub mnuLapBulananPNS_Click()
    frmLaporanBulananPegawaiC.Show
End Sub

Private Sub mnuLapHukuman_Click()
    frmLaporanHukuman.Show
End Sub

Private Sub mnuLapJmlPegawai_Click()
    frmLaporanBulananPegawai.Show
End Sub

Private Sub mnuLapKenaikanGajiBerkala_Click()
    frmLaporanRealisasiKenaikanGaji.Show
End Sub

Private Sub mnuLapKenaikanPangkat_Click()
    frmLaporanRealisasiKenaikanPangkat.Show
End Sub

Private Sub mnuLapNominatifPegawai_Click()
    frmLaporanRiwayatNOMINATIF.Show
End Sub

Private Sub mnuLaporanBukuCVPegawai_Click()
    strMenu = "CV Pegawai"
    frmKriteriaLaporan.Show
End Sub

Private Sub mnuLaporanBulananB_Click()
    frmLaporanBulananPegawaiB.Show
End Sub

Private Sub mnuLaporanDaftarPegawai_Click()
    strMenu = "Daftar Pegawai"
    frmKriteriaLaporan.Show
End Sub

Private Sub mnuLaporanIndexPegawai_Click()
    strMenu = "Index Pegawai"
    frmKriteriaLaporan.Show
End Sub

Private Sub mnuLaporanMutasiPegawai_Click()
    frmLaporanRiwayatMutasiPegawai.Show
End Sub

Private Sub mnuLaporanRiwayatPrestasi_Click()
    frmLaporanRiwayatPrestasi.Show
End Sub

Private Sub mnuLaporanRiwayatTugasBelajar_Click()
    frmLaporanRiwayatTugas.Show
End Sub

Private Sub mnuLapPenonaktifanPegawai_Click()
    frmLaporanRiwayatNonAktifPegawai.Show
End Sub

Private Sub mnuLapUsulanPengangkatanPNS_Click()
    frmLaporanRealisasiPengangkatanPNS.Show
End Sub

Private Sub mnuLapUsulanPensiun_Click()
    frmLaporanRiwayatUsulanPensiun.Show
End Sub

Private Sub mnuLapUsulanTaperum_Click()
    frmLaporanRiwayatUsulanTaperum.Show
End Sub

Private Sub mnuLapUsulanTPHLBerhenti_Click()
    frmLaporanRiwayatUsulanTPHLBerhenti.Show
End Sub

Private Sub mnuLapUsulATPHL_Click()
    frmLaporanRiwayatUsulanTPHL.Show
End Sub

Private Sub mnuLogOff_Click()
    Dim adoCommand As New ADODB.Command
    openConnection
    sepuh = True
    strQuery = "UPDATE Login SET Status = '0' " & _
    "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
    adoCommand.ActiveConnection = dbConn
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    adoCommand.Execute

    dTglLogout = Now
    Call subSp_HistoryLoginAplikasi("U")
    Unload Me
End Sub

Private Sub mnuMalam_Click()
    Call subIsiJadwalTetap("M")
End Sub

Private Sub mnuMstrGAP_Click()
frmGapKompetensiM.Show
End Sub

Private Sub mnuNilaiKinerja_Click()
    frmKinerja_Transaksi.Show
    frmKinerja_Transaksi.frmSasaran.Visible = True
    frmKinerja_Transaksi.statusForm = "NILAI"
    frmKinerja_Transaksi.Label3.Caption = "Bulan"
    frmKinerja_Transaksi.DTPicker2.Visible = True
    frmKinerja_Transaksi.DTPicker1.Visible = False
    frmKinerja_Transaksi.Label1.Caption = "Penilaian Kinerja"
    'frmKinerja_Transaksi.frmPenilaian.Visible = True
    frmKinerja_Transaksi.fgData.Rows = 1
    frmKinerja_Transaksi.fgData.Cols = 1
    frmKinerja_Transaksi.fgdata2.Rows = 1
    frmKinerja_Transaksi.fgdata2.Cols = 1
    frmKinerja_Transaksi.cmdCetak.Visible = True
End Sub

Private Sub mnuNotifikasi_Click()
    frmNotifikasi.Show
End Sub

'//yayang.agus 2014-08-22
Private Sub mnuReguler_Click()
    Call subIsiJadwalTetap("R")
End Sub

Private Sub mnuMasterDataInsentif_Click()
    frmMasterInsentif.Show
End Sub

Private Sub mnuMasterF_Click()
    frmMasterFungsionalPegawai.Show
End Sub

Private Sub mnuMasterFu_Click()
    frmMasterPelayananFungsional.Show
End Sub

Private Sub mnuMonitoringStokBarangNM_Click()
    frmMonitoringStokBarangNonMedis.Show
End Sub

Private Sub mnuPagi_Click()
    Call subIsiJadwalTetap("P")
End Sub

Private Sub mnuPangGol_Click()
    frmPangkatGolonganPegawai.Show
End Sub

Private Sub mnuPayroll_Click()
    frmLaporanGajiPegawai.Show
End Sub

Private Sub mnuPemakaianBahandanAlat_Click()
    frmPemakaianBahanAlat.Show
End Sub

Private Sub mnuPenilaianPegawai_Click()
    frmDataPerhitunganNilai.Show
End Sub

Private Sub mnuRekapKinerja_Click()
    frmRekapKinerja.Show
End Sub

Private Sub mnureqPNS_Click()
    On Error GoTo hell
    Dim pesan As VbMsgBoxResult
    strSQL = "SELECT DISTINCT * " & _
    " FROM V_RekapPegawaiBerdasarkanPendidikan "

    Call msubRecFO(rs, strSQL)
    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    frmCetakDaftarPegawaibyPendidikan.Show
    Exit Sub
hell:
End Sub

Private Sub mnuriwayatgaji_Click()
    frmLaporanGajiBerkala.Show
End Sub

Private Sub mnuRiwayatKirimdanTerimaTenagaMedis_Click()
    frmRiwayatKirimDanTerimaTenagaPemeriksa.Show
End Sub

Private Sub mnuriwayatpangkat_Click()
    frmLaporanRiwayatPangkat.Show
End Sub

Private Sub mnuriwayatstatus_Click()
    frmLaporanStatusCuti.Show
End Sub

Private Sub mnuSasaranKinerja_Click()
    frmKinerja_Transaksi.Show
    frmKinerja_Transaksi.frmSasaran.Visible = True
    frmKinerja_Transaksi.statusForm = "SASARAN"
    frmKinerja_Transaksi.Label3.Caption = "Tahun"
    frmKinerja_Transaksi.DTPicker2.Visible = False
    frmKinerja_Transaksi.DTPicker1.Visible = True
    frmKinerja_Transaksi.Label1.Caption = "Sasaran Per Tahun"
    frmKinerja_Transaksi.fgData.Rows = 1
    frmKinerja_Transaksi.fgData.Cols = 1
    frmKinerja_Transaksi.fgdata2.Rows = 1
    frmKinerja_Transaksi.fgdata2.Cols = 1
    frmKinerja_Transaksi.cmdCetak.Visible = False
End Sub

Private Sub mnuShift_Click()
    frmMasterAbsensi.Show
End Sub

Private Sub mnuSiang_Click()
    Call subIsiJadwalTetap("S")
End Sub

Private Sub mnuPersonalPegawai_Click()
    frmDataPegawaiNew.Show
End Sub

Private Sub mnuRiwayatExtraPelatihan_Click()
    frmRiwayatExtraPelatihan.Show
End Sub

Private Sub mnuRiwayatOrganisasi_Click()
    frmRiwayatOrganisasi.Show
End Sub

Private Sub mnuRiwayatPekerjaan_Click()
    frmRiwayatPekerjaan.Show
End Sub

Private Sub mnuRiwayatPendidikanFormal_Click()
    frmRiwayatPendidikanFormal.Show
End Sub

Private Sub mnuRiwayatPendidikanNonFormal_Click()
    frmRiwayatPendidikanNonFormal.Show
End Sub

Private Sub mnuRiwayatPerjalananDinas_Click()
    frmLaporanPerjalananDinas.Show
End Sub

Private Sub mnuRiwayatPrestasi_Click()
    frmRiwayatPrestasi.Show
End Sub

Private Sub mnuPIN_Click()
    frmRegistrasiUser_Offline.Show
End Sub

Private Sub mnuSelesai_Click()
    Dim pesan As VbMsgBoxResult
    Dim adoCommand As New ADODB.Command
    pesan = MsgBox("Tutup aplikasi ", vbQuestion + vbYesNo, "Konfirmasi")
    If pesan = vbYes Then

        openConnection
        strQuery = "UPDATE Login SET Status = '0' " & _
        "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        adoCommand.ActiveConnection = dbConn
        adoCommand.CommandText = strQuery
        adoCommand.CommandType = adCmdText
        adoCommand.Execute

        dTglLogout = Now
        Call subSp_HistoryLoginAplikasi("U")

        End
    Else
    End If
End Sub

Private Sub mnuSettingPrinter_Click()
    frmSetupPrinter2.Show
End Sub

Private Sub mnuTempatBertugas_Click()
    frmRiwayatTempatBertugas.Show
End Sub

Private Sub mnuStatusAbsen_Click()
    frmStatusAbsensi.Show
End Sub

Private Sub mnutanggal_Click()
    frmDataTanggal.Show
End Sub

Private Sub mnustatusabsensi_Click()
    frmStatusAbsensi.Show
End Sub

Private Sub mnuSubMasterAbsen_Click()
    frmMasterAbsensi.Show
End Sub

Private Sub mnuTentangMedifirst2000_Click()
    frmAbout.Show
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Exit Sub
    PopupMenu mnudata
End Sub

Private Sub mnuTitle_Click()
    frmTitle.Show
End Sub

Private Sub mnuTransaksiRiwayatUsulan_Click()
    frmUsulanRealisasi.Show
End Sub

Private Sub mnuTypePeg_Click()
    frmTypePegawai.Show
End Sub

Private Sub mnuUsulan_Click()
    frmFilterUsulan.Show
End Sub

Private Sub mPemesananBarang_Click()
    frmPemesananBarang.Show
End Sub

Private Sub mRekapitulasiTranBrgNM_Click()
    mstrKdKelompokBarang = "01"
    frmDataTransaksiBarang.Show
End Sub

Private Sub MS_Click()
    'frmMasterKinerja.Show
    frmKinerja_Kategory.Show
End Sub

Private Sub mStokBarang_Click()
    frmStokBrg.Show
End Sub

Private Sub mStokBarangNM_Click()
    frmStokBarangNonMedis.Show
End Sub

Private Sub PembayaranGajiPegawai_Click()
    frmPembayaranGajiPegawai2.Show
End Sub

Private Sub POAKaryawan_Click()
    frmDaftarPakaiAlkesKaryawan.Show
End Sub

Private Sub RiwayatGaji_Click()
    frmRiwayatGaji.Show
End Sub

Private Sub RekapLaporanGaji_Click()
    frmRekapLaporanGajiNew.Show
End Sub

Private Sub RiwayatPendidikandanLatihan_Click()
    frmRiwayatPendidikanDanPelatihanDiklat.Show
End Sub



Private Sub Timer1_Timer()
    Dim splt() As String
    
    LamaNotif = 1
    If GetSetting("SDM", "Notif", "Ultah") <> "" Then
        splt = Split(GetSetting("SDM", "Notif", "Ultah"), "~")
        If splt(0) = Date Then
            If splt(1) = "1" Then LamaNotif = 360
            If splt(1) = "5" Then LamaNotif = 360 * 5
            If splt(1) = "10" Then LamaNotif = 360 * 10
            If splt(1) = "0" Then LamaNotif = 0
        End If
    End If
    TimerNotif = TimerNotif + 1
    If TimerNotif = LamaNotif Then
        TimerNotif = 0
        'strSQL = "select * from datapegawai where " & _
                 "DatePart(Day, tgllahir) = DatePart(Day, GETDATE()) And DatePart(Month, tgllahir) = DatePart(Month, GETDATE())"
        strSQL = "select * from V_Notifikasi where tgl ='" & Format(Now(), "yyyy-MM-dd") & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            FrmNotifikasi2.Text1.Text = rs(1) & vbCrLf & rs(3)
            Beep
            FrmNotifikasi2.Show
        End If
    End If
End Sub




