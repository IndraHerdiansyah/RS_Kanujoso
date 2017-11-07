VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDRH 
   Caption         =   "frmCetakDaftarNoAbsensi"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakDRH.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakDRH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Report As New crDRH

Private Sub Form_Load()
On Error GoTo hell
Dim adocomd As New ADODB.Command
Set adocomd = Nothing
Me.WindowState = 2
adocomd.ActiveConnection = dbConn
    
   adocomd.CommandText = strSQL
   adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd
    With Report
'    .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
'    .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
'    .txtAlamatRS.SetText strNAlamatRS
'    .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos : " & " " & strNKodepos & " "

    .txtNamaLengkap.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
    .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
    .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat").Value), "", rs.Fields("Pangkat").Value)
    .txtTglLahir.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", Format(rs.Fields("TglLahir").Value, "dd MMMM yyyy"))
    .txtTempatLahir.SetText IIf(IsNull(rs.Fields("TempatLahir").Value), "", rs.Fields("TempatLahir").Value)
    If rs.Fields("JenisKelamin").Value = "L" Then
        .txtJK.SetText "Pria"
    ElseIf rs.Fields("JenisKelamin").Value = "P" Then
        .txtJK.SetText "Wanita"
    End If
    .txtUmur.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("Umur").Value)
    .txtAgama.SetText IIf(IsNull(rs.Fields("Agama").Value), "", rs.Fields("Agama").Value)
    .txtStatusKawin.SetText IIf(IsNull(rs.Fields("StatusPerkawinan").Value), "", rs.Fields("StatusPerkawinan").Value)
    .txtAlamat.SetText IIf(IsNull(rs.Fields("AlamatLengkap").Value), "", rs.Fields("AlamatLengkap").Value)
    .txtKelurahan.SetText IIf(IsNull(rs.Fields("Kelurahan").Value), "", rs.Fields("Kelurahan").Value)
    .txtKecamatan.SetText IIf(IsNull(rs.Fields("Kecamatan").Value), "", rs.Fields("Kecamatan").Value)
    .txtKota.SetText IIf(IsNull(rs.Fields("KotaKabupaten").Value), "", rs.Fields("KotaKabupaten").Value)
    .txtPropinsi.SetText IIf(IsNull(rs.Fields("Propinsi").Value), "", rs.Fields("Propinsi").Value)
    .txtTinggi.SetText IIf(IsNull(rs.Fields("TinggiBadan").Value), "", rs.Fields("TinggiBadan").Value)
    .txtBerat.SetText IIf(IsNull(rs.Fields("BeratBadan").Value), "", rs.Fields("BeratBadan").Value)
    .txtRambut.SetText IIf(IsNull(rs.Fields("JenisRambut").Value), "", rs.Fields("JenisRambut").Value)
    .txtMuka.SetText IIf(IsNull(rs.Fields("BentukMuka").Value), "", rs.Fields("BentukMuka").Value)
    .txtWarna.SetText IIf(IsNull(rs.Fields("WarnaKulit").Value), "", rs.Fields("WarnaKulit").Value)
    .txtCiri.SetText IIf(IsNull(rs.Fields("CiriCiriKhas").Value), "", rs.Fields("CiriCiriKhas").Value)
    .txtCacat.SetText IIf(IsNull(rs.Fields("CacatTubuh").Value), "", rs.Fields("CacatTubuh").Value)
    .txtHobi.SetText IIf(IsNull(rs.Fields("Hobby").Value), "", rs.Fields("Hobby").Value)
    
'CETAK LEMBAR KE 2 (RIWAYAT PENDIDIKAN FORMAL & NON)
    strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan='02' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtSD.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtSDJurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtSDThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtSDTempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtSDPimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
    End If

    strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan in('03','24') "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtSMP.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtSMPJurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtSMPThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtSMPTempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtSMPPimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
    End If
    
    strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan in ('04','22','23','24','27','29','30') "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtSMA.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtSMAJurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtSMAThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtSMATempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtSMAPimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
   End If
   strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan ='05' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtD1.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtD1Jurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtD1ThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtD1Tempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtD1Pimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
    End If
    strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan ='06' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtD2.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtD2Jurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtD2ThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtD2Tempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtD2Pimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
    End If
    strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan ='07' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtD3.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtD3Jurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtD3ThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtD3Tempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtD3Pimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
    End If
    strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan ='08' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtD4.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtD4Jurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtD4ThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtD4Tempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtD4Pimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
    End If
    strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan ='09' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtS1.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtS1Jurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtS1ThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtS1Tempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtS1Pimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
    End If
    strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan in('10','11') "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtS23.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtS23Jurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtS23ThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtS23Tempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtS23Pimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
    End If
    strSQL = "select * from V_CetakDRHPegawai where IdPegawai ='" & mstrIdPegawai & "' and KdPendidikan ='25' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtDokter.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtDokterJurusan.SetText IIf(IsNull(rs.Fields("FakultasJurusan").Value), "", rs.Fields("FakultasJurusan").Value)
        .Subreport1_txtDokterThnSTTB.SetText IIf(IsNull(rs.Fields("TahunSTTB").Value), "", rs.Fields("TahunSTTB").Value)
        .Subreport1_txtDokterTempat.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtDokterPimpinan.SetText IIf(IsNull(rs.Fields("PimpinanPendidikan").Value), "", rs.Fields("PimpinanPendidikan").Value)
    End If
    
    strSQL = "SELECT IdPegawai, NoUrut, NamaPendidikan, LamaPendidikan, YEAR(TglSertifikat) AS TahunSertifikat, AlamatPendidikan, Keterangan FROM RiwayatPendidikanNonFormal where IdPegawai ='" & mstrIdPegawai & "' and NoUrut ='001' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtPendidikanNF.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtLamanya.SetText IIf(IsNull(rs.Fields("LamaPendidikan").Value), "", rs.Fields("LamaPendidikan").Value)
        .Subreport1_txtTahunSertifikat.SetText IIf(IsNull(rs.Fields("TahunSertifikat").Value), "", rs.Fields("TahunSertifikat").Value)
        .Subreport1_txtAlamatNF.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtKeteranganNF.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    strSQL = "SELECT IdPegawai, NoUrut, NamaPendidikan, LamaPendidikan, YEAR(TglSertifikat) AS TahunSertifikat, AlamatPendidikan, Keterangan FROM RiwayatPendidikanNonFormal where IdPegawai ='" & mstrIdPegawai & "' and NoUrut ='002' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtPendidikanNF2.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtLamanya2.SetText IIf(IsNull(rs.Fields("LamaPendidikan").Value), "", rs.Fields("LamaPendidikan").Value)
        .Subreport1_txtTahunSertifikat2.SetText IIf(IsNull(rs.Fields("TahunSertifikat").Value), "", rs.Fields("TahunSertifikat").Value)
        .Subreport1_txtAlamatNF2.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtKeteranganNF2.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    strSQL = "SELECT IdPegawai, NoUrut, NamaPendidikan, LamaPendidikan, YEAR(TglSertifikat) AS TahunSertifikat, AlamatPendidikan, Keterangan FROM RiwayatPendidikanNonFormal where IdPegawai ='" & mstrIdPegawai & "' and NoUrut ='003' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtPendidikanNF3.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtLamanya3.SetText IIf(IsNull(rs.Fields("LamaPendidikan").Value), "", rs.Fields("LamaPendidikan").Value)
        .Subreport1_txtTahunSertifikat3.SetText IIf(IsNull(rs.Fields("TahunSertifikat").Value), "", rs.Fields("TahunSertifikat").Value)
        .Subreport1_txtAlamatNF3.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtKeteranganNF3.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    strSQL = "SELECT IdPegawai, NoUrut, NamaPendidikan, LamaPendidikan, YEAR(TglSertifikat) AS TahunSertifikat, AlamatPendidikan, Keterangan FROM RiwayatPendidikanNonFormal where IdPegawai ='" & mstrIdPegawai & "' and NoUrut ='004' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport1_txtPendidikanNF4.SetText rs.Fields("NamaPendidikan").Value
        .Subreport1_txtLamanya4.SetText IIf(IsNull(rs.Fields("LamaPendidikan").Value), "", rs.Fields("LamaPendidikan").Value)
        .Subreport1_txtTahunSertifikat4.SetText IIf(IsNull(rs.Fields("TahunSertifikat").Value), "", rs.Fields("TahunSertifikat").Value)
        .Subreport1_txtAlamatNF4.SetText IIf(IsNull(rs.Fields("AlamatPendidikan").Value), "", rs.Fields("AlamatPendidikan").Value)
        .Subreport1_txtKeteranganNF4.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
'CETAK LEMBAR KE 3 (RIWAYAT PANGKAT GOLONGAN GAJI)
    strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='01' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat1.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol1.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji1.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD1.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK1.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK1.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
    
    strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='02' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat2.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol2.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji2.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD2.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK2.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK2.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
     strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='03' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat3.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol3.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji3.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD3.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK3.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK3.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
     strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='04' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat4.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol4.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji4.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD4.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK4.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK4.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
     strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='05' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat5.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol5.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji5.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD5.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK5.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK5.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
     strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='06' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat6.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol6.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji6.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD6.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK6.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK6.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
     strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='07' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat7.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol7.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji7.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD7.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK7.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK7.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
     strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='08' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat8.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol8.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji8.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD8.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK8.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK8.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
     strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='09' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat9.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol9.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji9.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD9.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK9.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK9.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
    strSQL = "select * from V_CetakDRHPangkatGolGaji where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='10' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtPangkat10.SetText rs.Fields("NamaPangkat").Value
        .Subreport2_txtGol10.SetText rs.Fields("NamaGolongan").Value
        .Subreport2_txtGaji10.SetText IIf(IsNull(rs.Fields("Jumlah").Value), "", rs.Fields("Jumlah").Value)
        .Subreport2_txtTTD10.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSK10.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSK10.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
    
    strSQL = "select IdPegawai, NoUrut, NamaPerusahaan, CONVERT(varchar, TglMulai, 103) + ' ' + 's/d' + ' ' + CONVERT(varchar, TglAkhir, 103) AS Periode, JabatanPosisi, GajiPokok, TandaTanganSK, NoSK, TglSK FROM RiwayatPekerjaan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='01' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtNamaKerja1.SetText rs.Fields("NamaPerusahaan").Value
        .Subreport2_txtKerjaTgl1.SetText IIf(IsNull(rs.Fields("Periode").Value), "", rs.Fields("Periode").Value)
        .Subreport2_txtJabatanKerja1.SetText IIf(IsNull(rs.Fields("JabatanPosisi").Value), "", rs.Fields("JabatanPosisi").Value)
        .Subreport2_txtGajiKerja1.SetText IIf(IsNull(rs.Fields("GajiPokok").Value), "", rs.Fields("GajiPokok").Value)
        .Subreport2_txtTTDSKKerja1.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSKKerja1.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSKKerja1.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPerusahaan, CONVERT(varchar, TglMulai, 103) + ' ' + 's/d' + ' ' + CONVERT(varchar, TglAkhir, 103) AS Periode, JabatanPosisi, GajiPokok, TandaTanganSK, NoSK, TglSK FROM RiwayatPekerjaan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='02' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtNamaKerja2.SetText rs.Fields("NamaPerusahaan").Value
        .Subreport2_txtKerjaTgl2.SetText IIf(IsNull(rs.Fields("Periode").Value), "", rs.Fields("Periode").Value)
        .Subreport2_txtJabatanKerja2.SetText IIf(IsNull(rs.Fields("JabatanPosisi").Value), "", rs.Fields("JabatanPosisi").Value)
        .Subreport2_txtGajiKerja2.SetText IIf(IsNull(rs.Fields("GajiPokok").Value), "", rs.Fields("GajiPokok").Value)
        .Subreport2_txtTTDSKKerja2.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSKKerja2.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSKKerja2.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPerusahaan, CONVERT(varchar, TglMulai, 103) + ' ' + 's/d' + ' ' + CONVERT(varchar, TglAkhir, 103) AS Periode, JabatanPosisi, GajiPokok, TandaTanganSK, NoSK, TglSK FROM RiwayatPekerjaan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='03' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtNamaKerja3.SetText rs.Fields("NamaPerusahaan").Value
        .Subreport2_txtKerjaTgl3.SetText IIf(IsNull(rs.Fields("Periode").Value), "", rs.Fields("Periode").Value)
        .Subreport2_txtJabatanKerja3.SetText IIf(IsNull(rs.Fields("JabatanPosisi").Value), "", rs.Fields("JabatanPosisi").Value)
        .Subreport2_txtGajiKerja3.SetText IIf(IsNull(rs.Fields("GajiPokok").Value), "", rs.Fields("GajiPokok").Value)
        .Subreport2_txtTTDSKKerja3.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSKKerja3.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSKKerja3.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPerusahaan, CONVERT(varchar, TglMulai, 103) + ' ' + 's/d' + ' ' + CONVERT(varchar, TglAkhir, 103) AS Periode, JabatanPosisi, GajiPokok, TandaTanganSK, NoSK, TglSK FROM RiwayatPekerjaan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='04' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtNamaKerja4.SetText rs.Fields("NamaPerusahaan").Value
        .Subreport2_txtKerjaTgl4.SetText IIf(IsNull(rs.Fields("Periode").Value), "", rs.Fields("Periode").Value)
        .Subreport2_txtJabatanKerja4.SetText IIf(IsNull(rs.Fields("JabatanPosisi").Value), "", rs.Fields("JabatanPosisi").Value)
        .Subreport2_txtGajiKerja4.SetText IIf(IsNull(rs.Fields("GajiPokok").Value), "", rs.Fields("GajiPokok").Value)
        .Subreport2_txtTTDSKKerja4.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSKKerja4.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSKKerja4.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPerusahaan, CONVERT(varchar, TglMulai, 103) + ' ' + 's/d' + ' ' + CONVERT(varchar, TglAkhir, 103) AS Periode, JabatanPosisi, GajiPokok, TandaTanganSK, NoSK, TglSK FROM RiwayatPekerjaan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut='05' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport2_txtNamaKerja5.SetText rs.Fields("NamaPerusahaan").Value
        .Subreport2_txtKerjaTgl5.SetText IIf(IsNull(rs.Fields("Periode").Value), "", rs.Fields("Periode").Value)
        .Subreport2_txtJabatanKerja5.SetText IIf(IsNull(rs.Fields("JabatanPosisi").Value), "", rs.Fields("JabatanPosisi").Value)
        .Subreport2_txtGajiKerja5.SetText IIf(IsNull(rs.Fields("GajiPokok").Value), "", rs.Fields("GajiPokok").Value)
        .Subreport2_txtTTDSKKerja5.SetText IIf(IsNull(rs.Fields("TandaTanganSK").Value), "", rs.Fields("TandaTanganSK").Value)
        .Subreport2_txtNoSKKerja5.SetText IIf(IsNull(rs.Fields("NoSK").Value), "", rs.Fields("NoSK").Value)
        .Subreport2_txtTglSKKerja5.SetText IIf(IsNull(rs.Fields("TglSK").Value), "", rs.Fields("TglSK").Value)
    End If
    
'CETAK LEMBAR 4 (RIWAYAT KELUARGA)
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('01','02') "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport3_txtIstri1.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtTempatLahirIstri1.SetText IIf(IsNull(rs.Fields("TempatLahir").Value), "", rs.Fields("TempatLahir").Value)
        .Subreport3_txtTglLahirIstri1.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtTglNikah1.SetText IIf(IsNull(rs.Fields("TglKawin").Value), "", rs.Fields("TglKawin").Value)
        .Subreport3_txtPekerjaanIstri1.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetIstri1.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('03','32','33') "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strNoUrut = rs.Fields("NoUrut").Value
        .Subreport3_txtAnak1.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtTmpLahirAnak1.SetText IIf(IsNull(rs.Fields("TempatLahir").Value), "", rs.Fields("TempatLahir").Value)
        .Subreport3_txtTglLahirAnak1.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtJKAnak1.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport3_txtPekerjaanAnak1.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetAnak1.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('03','32','33') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strNoUrut = rs.Fields("NoUrut").Value
        .Subreport3_txtAnak2.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtTmpLahirAnak2.SetText IIf(IsNull(rs.Fields("TempatLahir").Value), "", rs.Fields("TempatLahir").Value)
        .Subreport3_txtTglLahirAnak2.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtJKAnak2.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport3_txtPekerjaanAnak2.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetAnak2.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('03','32','33') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strNoUrut = rs.Fields("NoUrut").Value
        .Subreport3_txtAnak3.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtTmpLahirAnak3.SetText IIf(IsNull(rs.Fields("TempatLahir").Value), "", rs.Fields("TempatLahir").Value)
        .Subreport3_txtTglLahirAnak3.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtJKAnak3.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport3_txtPekerjaanAnak3.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetAnak3.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('03','32','33') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strNoUrut = rs.Fields("NoUrut").Value
        .Subreport3_txtAnak4.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtTmpLahirAnak4.SetText IIf(IsNull(rs.Fields("TempatLahir").Value), "", rs.Fields("TempatLahir").Value)
        .Subreport3_txtTglLahirAnak4.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtJKAnak4.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport3_txtPekerjaanAnak4.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetAnak4.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('03','32','33') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strNoUrut = rs.Fields("NoUrut").Value
        .Subreport3_txtAnak5.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtTmpLahirAnak5.SetText IIf(IsNull(rs.Fields("TempatLahir").Value), "", rs.Fields("TempatLahir").Value)
        .Subreport3_txtTglLahirAnak5.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtJKAnak5.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport3_txtPekerjaanAnak5.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetAnak5.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('06','34') "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        
        .Subreport3_txtAyah.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtAyahLahir.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtKerjaAyah.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetAyah.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan ='05' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        
        .Subreport3_txtIbu.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtIbuLahir.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtKerjaIbu.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetIbu.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan ='12' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strNoUrut = rs.Fields("NoUrut").Value
        .Subreport3_txtMertua1.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtMertuaLahir1.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtMertuaKerja1.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetMertua1.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan ='12' and NoUrut= '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
       
        .Subreport3_txtMertua2.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtMertuaLahir2.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtMertuaKerja2.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetMertua2.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan ='12' and NoUrut= '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        
        .Subreport3_txtMertua3.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtMertuaLahir3.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtMertuaKerja3.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetMertua3.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan ='12' and NoUrut= '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        
        .Subreport3_txtMertua4.SetText rs.Fields("NamaLengkap").Value
        .Subreport3_txtMertuaLahir4.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport3_txtMertuaKerja4.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport3_txtKetMertua4.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
'CETAK LEMBAR 4 (RIWAYAT KELUARGA kANDUNG)
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('14','18') "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strNoUrut = rs.Fields("NoUrut").Value
        .Subreport4_txtkandung1.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtKandungLahir1.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKandung1.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKandungKerja1.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetKandung1.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('14','18') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport4_txtkandung2.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtKandungLahir2.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKandung2.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKandungKerja2.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetKandung2.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('14','18') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport4_txtkandung3.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtKandungLahir3.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKandung3.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKandungKerja3.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetKandung3.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('14','18') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport4_txtkandung4.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtKandungLahir4.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKandung4.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKandungKerja4.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetKandung4.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('14','18') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport4_txtkandung5.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtKandungLahir5.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKandung5.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKandungKerja5.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetKandung5.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('25','26') "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strNoUrut = rs.Fields("NoUrut").Value
        .Subreport4_txtIpar1.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtLahirIpar1.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKIpar1.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKerjaIpar1.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetIpar1.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('25','26') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport4_txtIpar2.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtLahirIpar2.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKIpar2.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKerjaIpar2.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetIpar2.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('25','26') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport4_txtIpar3.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtLahirIpar3.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKIpar3.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKerjaIpar3.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetIpar3.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('25','26') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport4_txtIpar4.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtLahirIpar4.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKIpar4.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKerjaIpar4.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetIpar4.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    If Left(strNoUrut, 1) = 1 Then
        strNoUrut = strNoUrut + 1
    ElseIf Left(strNoUrut, 1) = 0 Then
        strNoUrut = strNoUrut + 1
        strNoUrut = "0" + strNoUrut
    End If
    strSQL = "select * from V_CetakDRHRiwayatKeluarga where IdPegawai ='" & mstrIdPegawai & "' and KdHubungan in ('25','26') and NoUrut = '" & strNoUrut & "' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport4_txtIpar5.SetText rs.Fields("NamaLengkap").Value
        .Subreport4_txtLahirIpar5.SetText IIf(IsNull(rs.Fields("TglLahir").Value), "", rs.Fields("TglLahir").Value)
        .Subreport4_txtJKIpar5.SetText IIf(IsNull(rs.Fields("JenisKelamin").Value), "", rs.Fields("JenisKelamin").Value)
        .Subreport4_txtKerjaIpar5.SetText IIf(IsNull(rs.Fields("Pekerjaan").Value), "", rs.Fields("Pekerjaan").Value)
        .Subreport4_txtKetIpar5.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
    End If
    
'CETAK LEMBAR 5 (Riwayat Prestasi)
    strSQL = "select IdPegawai, NoUrut, NamaPenghargaan, YEAR(TglDiperoleh) AS Tahun, InstansiPemberi FROM  dbo.RiwayatPrestasi where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '01' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtBintang1.SetText rs.Fields("NamaPenghargaan").Value
        .Subreport5_txtThIns1.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport5_txtNegaraIns1.SetText IIf(IsNull(rs.Fields("InstansiPemberi").Value), "", rs.Fields("InstansiPemberi").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPenghargaan, YEAR(TglDiperoleh) AS Tahun, InstansiPemberi FROM  dbo.RiwayatPrestasi where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '02' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtBintang2.SetText rs.Fields("NamaPenghargaan").Value
        .Subreport5_txtThIns2.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport5_txtNegaraIns2.SetText IIf(IsNull(rs.Fields("InstansiPemberi").Value), "", rs.Fields("InstansiPemberi").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPenghargaan, YEAR(TglDiperoleh) AS Tahun, InstansiPemberi FROM  dbo.RiwayatPrestasi where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '03' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtBintang3.SetText rs.Fields("NamaPenghargaan").Value
        .Subreport5_txtThIns3.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport5_txtNegaraIns3.SetText IIf(IsNull(rs.Fields("InstansiPemberi").Value), "", rs.Fields("InstansiPemberi").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPenghargaan, YEAR(TglDiperoleh) AS Tahun, InstansiPemberi FROM  dbo.RiwayatPrestasi where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '04' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtBintang4.SetText rs.Fields("NamaPenghargaan").Value
        .Subreport5_txtThIns4.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport5_txtNegaraIns4.SetText IIf(IsNull(rs.Fields("InstansiPemberi").Value), "", rs.Fields("InstansiPemberi").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPenghargaan, YEAR(TglDiperoleh) AS Tahun, InstansiPemberi FROM  dbo.RiwayatPrestasi where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '05' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtBintang5.SetText rs.Fields("NamaPenghargaan").Value
        .Subreport5_txtThIns5.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport5_txtNegaraIns5.SetText IIf(IsNull(rs.Fields("InstansiPemberi").Value), "", rs.Fields("InstansiPemberi").Value)
    End If
    
    strSQL = "select IdPegawai, NoUrut, NegaraTujuan, TujuanKunjungan, PenyandangDana, CAST(DATEDIFF(day, TglPergi, TglPulang) AS varchar) + ' ' + 'hari' AS Lamanya FROM dbo.RiwayatPerjalananDinas where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '01' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtDinas1.SetText rs.Fields("NegaraTujuan").Value
        .Subreport5_txtTujuanDinas1.SetText IIf(IsNull(rs.Fields("TujuanKunjungan").Value), "", rs.Fields("TujuanKunjungan").Value)
        .Subreport5_txtLamaDinas1.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", rs.Fields("Lamanya").Value)
        .Subreport5_txtDana1.SetText IIf(IsNull(rs.Fields("PenyandangDana").Value), "", rs.Fields("PenyandangDana").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NegaraTujuan, TujuanKunjungan, PenyandangDana, CAST(DATEDIFF(day, TglPergi, TglPulang) AS varchar) + ' ' + 'hari' AS Lamanya FROM dbo.RiwayatPerjalananDinas where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '02' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtDinas2.SetText rs.Fields("NegaraTujuan").Value
        .Subreport5_txtTujuanDinas2.SetText IIf(IsNull(rs.Fields("TujuanKunjungan").Value), "", rs.Fields("TujuanKunjungan").Value)
        .Subreport5_txtLamaDinas2.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", rs.Fields("Lamanya").Value)
        .Subreport5_txtDana2.SetText IIf(IsNull(rs.Fields("PenyandangDana").Value), "", rs.Fields("PenyandangDana").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NegaraTujuan, TujuanKunjungan, PenyandangDana, CAST(DATEDIFF(day, TglPergi, TglPulang) AS varchar) + ' ' + 'hari' AS Lamanya FROM dbo.RiwayatPerjalananDinas where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '03' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtDinas3.SetText rs.Fields("NegaraTujuan").Value
        .Subreport5_txtTujuanDinas3.SetText IIf(IsNull(rs.Fields("TujuanKunjungan").Value), "", rs.Fields("TujuanKunjungan").Value)
        .Subreport5_txtLamaDinas3.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", rs.Fields("Lamanya").Value)
        .Subreport5_txtDana3.SetText IIf(IsNull(rs.Fields("PenyandangDana").Value), "", rs.Fields("PenyandangDana").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NegaraTujuan, TujuanKunjungan, PenyandangDana, CAST(DATEDIFF(day, TglPergi, TglPulang) AS varchar) + ' ' + 'hari' AS Lamanya FROM dbo.RiwayatPerjalananDinas where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '04' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtDinas4.SetText rs.Fields("NegaraTujuan").Value
        .Subreport5_txtTujuanDinas4.SetText IIf(IsNull(rs.Fields("TujuanKunjungan").Value), "", rs.Fields("TujuanKunjungan").Value)
        .Subreport5_txtLamaDinas4.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", rs.Fields("Lamanya").Value)
        .Subreport5_txtDana4.SetText IIf(IsNull(rs.Fields("PenyandangDana").Value), "", rs.Fields("PenyandangDana").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NegaraTujuan, TujuanKunjungan, PenyandangDana, CAST(DATEDIFF(day, TglPergi, TglPulang) AS varchar) + ' ' + 'hari' AS Lamanya FROM dbo.RiwayatPerjalananDinas where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '05' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport5_txtDinas5.SetText rs.Fields("NegaraTujuan").Value
        .Subreport5_txtTujuanDinas5.SetText IIf(IsNull(rs.Fields("TujuanKunjungan").Value), "", rs.Fields("TujuanKunjungan").Value)
        .Subreport5_txtLamaDinas5.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", rs.Fields("Lamanya").Value)
        .Subreport5_txtDana5.SetText IIf(IsNull(rs.Fields("PenyandangDana").Value), "", rs.Fields("PenyandangDana").Value)
    End If

    Call CetakPelatihan
    
'        .SelectPrinter sDriver, sPrinter, vbNull
'         settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    
    End With
    

    
    
    With CRViewer1
        .ReportSource = Report
        .EnableGroupTree = False
        .EnableExportButton = True
        .ViewReport
        .Zoom 100
    End With
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub CetakPelatihan()
'CETAK LEMBAR 6 (Riwayat Extra Pelatihan)
With Report
    strSQL = "select IdPegawai, NoUrut, NamaPelatihan, KedudukanPeranan, CAST(MONTH(TglMulai) AS varchar) + '' + '/' + '' + CAST(YEAR(TglMulai) AS varchar) AS Tahun, InstansiPenyelenggara, AlamatPenyelenggara FROM  dbo.RiwayatExtraPelatihan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '001' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport6_txtLatih1.SetText rs.Fields("NamaPelatihan").Value
        .Subreport6_txtPeran1.SetText IIf(IsNull(rs.Fields("KedudukanPeranan").Value), "", rs.Fields("KedudukanPeranan").Value)
        .Subreport6_txtTahun1.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport6_txtInstansi1.SetText IIf(IsNull(rs.Fields("InstansiPenyelenggara").Value), "", rs.Fields("InstansiPenyelenggara").Value)
        .Subreport6_txtTempatLatih1.SetText IIf(IsNull(rs.Fields("AlamatPenyelenggara").Value), "", rs.Fields("AlamatPenyelenggara").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPelatihan, KedudukanPeranan, CAST(MONTH(TglMulai) AS varchar) + '' + '/' + '' + CAST(YEAR(TglMulai) AS varchar) AS Tahun, InstansiPenyelenggara, AlamatPenyelenggara FROM  dbo.RiwayatExtraPelatihan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '002' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport6_txtLatih2.SetText rs.Fields("NamaPelatihan").Value
        .Subreport6_txtPeran2.SetText IIf(IsNull(rs.Fields("KedudukanPeranan").Value), "", rs.Fields("KedudukanPeranan").Value)
        .Subreport6_txtTahun2.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport6_txtInstansi2.SetText IIf(IsNull(rs.Fields("InstansiPenyelenggara").Value), "", rs.Fields("InstansiPenyelenggara").Value)
        .Subreport6_txtTempatLatih2.SetText IIf(IsNull(rs.Fields("AlamatPenyelenggara").Value), "", rs.Fields("AlamatPenyelenggara").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPelatihan, KedudukanPeranan, CAST(MONTH(TglMulai) AS varchar) + '' + '/' + '' + CAST(YEAR(TglMulai) AS varchar) AS Tahun, InstansiPenyelenggara, AlamatPenyelenggara FROM  dbo.RiwayatExtraPelatihan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '003' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport6_txtLatih3.SetText rs.Fields("NamaPelatihan").Value
        .Subreport6_txtPeran3.SetText IIf(IsNull(rs.Fields("KedudukanPeranan").Value), "", rs.Fields("KedudukanPeranan").Value)
        .Subreport6_txtTahun3.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport6_txtInstansi3.SetText IIf(IsNull(rs.Fields("InstansiPenyelenggara").Value), "", rs.Fields("InstansiPenyelenggara").Value)
        .Subreport6_txtTempatLatih3.SetText IIf(IsNull(rs.Fields("AlamatPenyelenggara").Value), "", rs.Fields("AlamatPenyelenggara").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPelatihan, KedudukanPeranan, CAST(MONTH(TglMulai) AS varchar) + '' + '/' + '' + CAST(YEAR(TglMulai) AS varchar) AS Tahun, InstansiPenyelenggara, AlamatPenyelenggara FROM  dbo.RiwayatExtraPelatihan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '004' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport6_txtLatih4.SetText rs.Fields("NamaPelatihan").Value
        .Subreport6_txtPeran4.SetText IIf(IsNull(rs.Fields("KedudukanPeranan").Value), "", rs.Fields("KedudukanPeranan").Value)
        .Subreport6_txtTahun4.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport6_txtInstansi4.SetText IIf(IsNull(rs.Fields("InstansiPenyelenggara").Value), "", rs.Fields("InstansiPenyelenggara").Value)
        .Subreport6_txtTempatLatih4.SetText IIf(IsNull(rs.Fields("AlamatPenyelenggara").Value), "", rs.Fields("AlamatPenyelenggara").Value)
    End If
    strSQL = "select IdPegawai, NoUrut, NamaPelatihan, KedudukanPeranan, CAST(MONTH(TglMulai) AS varchar) + '' + '/' + '' + CAST(YEAR(TglMulai) AS varchar) AS Tahun, InstansiPenyelenggara, AlamatPenyelenggara FROM  dbo.RiwayatExtraPelatihan where IdPegawai ='" & mstrIdPegawai & "' and NoUrut = '005' "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        .Subreport6_txtLatih5.SetText rs.Fields("NamaPelatihan").Value
        .Subreport6_txtPeran5.SetText IIf(IsNull(rs.Fields("KedudukanPeranan").Value), "", rs.Fields("KedudukanPeranan").Value)
        .Subreport6_txtTahun5.SetText IIf(IsNull(rs.Fields("Tahun").Value), "", rs.Fields("Tahun").Value)
        .Subreport6_txtInstansi5.SetText IIf(IsNull(rs.Fields("InstansiPenyelenggara").Value), "", rs.Fields("InstansiPenyelenggara").Value)
        .Subreport6_txtTempatLatih5.SetText IIf(IsNull(rs.Fields("AlamatPenyelenggara").Value), "", rs.Fields("AlamatPenyelenggara").Value)
    End If
End With
Report.Subreport8_txtPembuat.SetText mstrNama
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmCetakDRH = Nothing
End Sub

