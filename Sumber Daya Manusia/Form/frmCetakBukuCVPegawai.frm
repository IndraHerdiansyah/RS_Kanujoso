VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakBukuCVPegawai 
   Caption         =   "CETAK"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakBukuCVPegawai.frx":0000
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
Attribute VB_Name = "frmCetakBukuCVPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crCVPegawai


Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Dim laporan As New ADODB.Command
    Dim laporan2 As New ADODB.Command
    Dim laporan3 As New ADODB.Command
    Dim laporan4 As New ADODB.Command
    Dim laporan5 As New ADODB.Command
    Dim laporan6 As New ADODB.Command
    Dim laporan7 As New ADODB.Command
    Dim laporan8 As New ADODB.Command
    Dim laporan9 As New ADODB.Command
    Dim laporan10 As New ADODB.Command
    Dim laporan11 As New ADODB.Command
    Dim laporan12 As New ADODB.Command
    Dim laporan13 As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    With laporan
        .ActiveConnection = dbConn
        .CommandText = "SELECT * from V_S_DataPegawai WHERE IdPegawai = '" & mstrIdPegawai & "'"
        .CommandType = adCmdText
    End With

'    With laporan02
'        .ActiveConnection = dbConn
'        .CommandText = "SELECT * from V_TempatBertugas WHERE IdPegawai = '" & frmKriteriaLaporan.txtIdAwal.Text & "'"
'        .CommandType = adCmdText
'    End With

    With laporan2
        .ActiveConnection = dbConn
        .CommandText = "SELECT * from DetailPegawai WHERE IdPegawai = '" & mstrIdPegawai & "'"
        .CommandType = adCmdText
    End With

    With laporan3
        .ActiveConnection = dbConn
        .CommandText = "SELECT * from DataAlamatPegawai WHERE IdPegawai = '" & mstrIdPegawai & "'"
        .CommandType = adCmdText
    End With

'    With laporan12
'        .ActiveConnection = dbConn
'        .CommandText = "SELECT * from V_KeluargaPegawai WHERE IdPegawai = '" & frmKriteriaLaporan.txtIdAwal.Text & "'"
'        .CommandType = adCmdText
'    End With

    With laporan4
        .ActiveConnection = dbConn
        .CommandText = "SELECT * from RiwayatPendidikanFormal WHERE IdPegawai = '" & mstrIdPegawai & "'"
        .CommandType = adCmdText
    End With

    With laporan5
        .ActiveConnection = dbConn
        .CommandText = "SELECT * from RiwayatPendidikanNonFormal WHERE IdPegawai = '" & mstrIdPegawai & "'"
        .CommandType = adCmdText
    End With

    With laporan6
        .ActiveConnection = dbConn
        .CommandText = "SELECT * from RiwayatExtraPelatihan WHERE IdPegawai = '" & mstrIdPegawai & "'"
        .CommandType = adCmdText
    End With

    With laporan9
        .ActiveConnection = dbConn
        .CommandText = "SELECT * from RiwayatOrganisasi WHERE IdPegawai = '" & mstrIdPegawai & "'"
        .CommandType = adCmdText
    End With

    With laporan10
        .ActiveConnection = dbConn
        .CommandText = "SELECT * from RiwayatPekerjaan WHERE IdPegawai = '" & mstrIdPegawai & "'"
        .CommandType = adCmdText
    End With

'    With laporan11
'        .ActiveConnection = dbConn
'        .CommandText = "SELECT * from RiwayatPerjalananDinas WHERE IdPegawai = '" & frmKriteriaLaporan.txtIdAwal.Text & "'"
'        .CommandType = adCmdText
'    End With

    With laporan12
        .ActiveConnection = dbConn
        .CommandText = "SELECT * from RiwayatPrestasi WHERE IdPegawai = '" & mstrIdPegawai & "'"
        .CommandType = adCmdText
    End With

'    With laporan13
'        .ActiveConnection = dbConn
'        .CommandText = "SELECT * from V_RiwayatGaji WHERE IdPegawai = '" & frmKriteriaLaporan.txtIdAwal.Text & "'"
'        .CommandType = adCmdText
'    End With

    With Report
        .Database.AddADOCommand dbConn, laporan
        .txtNamaRS.SetText strNNamaRS '& " " & strkelasRS & " " & strketkelasRS
        .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
        .txtAlamatRS.SetText strNAlamatRS
        .txtInstalasi.SetText "HUMAN RESOURCE DEPARTMENT"
        .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos : " & " " & strNKodepos & " "
        .usIdPegawai.SetUnboundFieldSource ("{ado.IdPegawai}")
        .usJenisPegawai.SetUnboundFieldSource ("{ado.JenisPegawai}")
        .usNamaLengkap.SetUnboundFieldSource ("{ado.NamaLengkap}")
        .usJenisKelamin.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .usTempatLahir.SetUnboundFieldSource ("{ado.TempatLahir}")
        .udTglLahir.SetUnboundFieldSource ("{ado.TglLahir}")
        .usPangkat.SetUnboundFieldSource ("{ado.NamaPangkat}")
        .usGolongan.SetUnboundFieldSource ("{ado.NamaGolongan}")
        .usJabatan.SetUnboundFieldSource ("{ado.NamaJabatan}")
        .usPendidikanTerakhir.SetUnboundFieldSource ("{ado.Pendidikan}")
        .usNIP.SetUnboundFieldSource ("{ado.NIP}")

'        .Subreport1.OpenSubreport.Database.AddADOCommand dbConn, laporan2
'        .Subreport1_usRuangan.SetUnboundFieldSource ("{ado.NamaRuangan}")
'        .Subreport1_usTglMulai.SetUnboundFieldSource ("{ado.TglMulai}")
'        .Subreport1_usTglAkhir.SetUnboundFieldSource ("{ado.TglAkhir}")
'        .Subreport1_usNoSuratKeputusan.SetUnboundFieldSource ("{ado.NoSuratKeputusan}")

'        .Subreport2.OpenSubreport.Database.AddADOCommand dbConn, laporan2
'        .Subreport2_usAgama.SetUnboundFieldSource ("{ado.Agama}")
'        .Subreport2_usStatusPerkawinan.SetUnboundFieldSource ("{ado.StatusPerkawinan}")
'        .Subreport2_usGolonganDarah.SetUnboundFieldSource ("{ado.GolonganDarah}")
'        .Subreport2_usHobby.SetUnboundFieldSource ("{ado.Hobby}")
'        .Subreport2_usTinggiBadan.SetUnboundFieldSource ("{ado.TinggiBadan}")
'        .Subreport2_usBeratBadan.SetUnboundFieldSource ("{ado.BeratBadan}")
'        .Subreport2_usJenisRambut.SetUnboundFieldSource ("{ado.JenisRambut}")
'        .Subreport2_usBEntukMuka.SetUnboundFieldSource ("{ado.BentukMuka}")
'        .Subreport2_usWarnaKulit.SetUnboundFieldSource ("{ado.WarnaKulit}")
'        .Subreport2_usCiriCiriKhas.SetUnboundFieldSource ("{ado.CiriCiriKhas}")
'        .Subreport2_usCacatTubuh.SetUnboundFieldSource ("{ado.CacatTubuh}")

        .Subreport3.OpenSubreport.Database.AddADOCommand dbConn, laporan3
        .Subreport3_usAlamatLengkap.SetUnboundFieldSource ("{ado.AlamatLengkap}")
        .Subreport3_usKelurahan.SetUnboundFieldSource ("{ado.Kelurahan}")
        .Subreport3_usKecamatan.SetUnboundFieldSource ("{ado.Kecamatan}")
        .Subreport3_usKotaKabupaten.SetUnboundFieldSource ("{ado.KotaKabupaten}")
        .Subreport3_usPropinsi.SetUnboundFieldSource ("{ado.Propinsi}")
        .Subreport3_usRTRW.SetUnboundFieldSource ("{ado.RTRW}")
        .Subreport3_usKodePos.SetUnboundFieldSource ("{ado.KodePos}")
        .Subreport3_usTelepon.SetUnboundFieldSource ("{ado.Telepon}")

'        .Subreport4.OpenSubreport.Database.AddADOCommand dbConn, laporan5
'        .Subreport4_usNoUrutKeluarga.SetUnboundFieldSource ("{ado.NoUrut}")
'        .Subreport4_usHubunganKeluarga.SetUnboundFieldSource ("{ado.NamaHubungan}")
'        .Subreport4_usNamaLengkapKeluarga.SetUnboundFieldSource ("{ado.NamaLengkap}")
'        .Subreport4_usJenisKelaminKeluarga.SetUnboundFieldSource ("{ado.JenisKelamin}")
'        .Subreport4_usTglLahirKeluarga.SetUnboundFieldSource ("{ado.TglLahir}")
'        .Subreport4_usPekerjaanKeluarga.SetUnboundFieldSource ("{ado.Pekerjaan}")
'        .Subreport4_usPendidikanKeluarga.SetUnboundFieldSource ("{ado.Pendidikan}")
'        .Subreport4_usKeteranganKeluarga.SetUnboundFieldSource ("{ado.Keterangan}")

        'Untuk repot Pendidikan Formal
        .Subreport5.OpenSubreport.Database.AddADOCommand dbConn, laporan4
        .Subreport5_usNoUrutPF.SetUnboundFieldSource ("{ado.NoUrut}")
        .Subreport5_usNamaPendidikanPF.SetUnboundFieldSource ("{ado.NamaPendidikan}")
        .Subreport5_usJurusanPF.SetUnboundFieldSource ("{ado.FakultasJurusan}")
        .Subreport5_usTahunMasukPF.SetUnboundFieldSource ("{ado.TglMasuk}")
        .Subreport5_usTahunLulusPF.SetUnboundFieldSource ("{ado.TglLulus}")
        .Subreport5_usIPKPF.SetUnboundFieldSource ("{ado.IPK}")
        .Subreport5_usGradePF.SetUnboundFieldSource ("{ado.GradeKelulusan}")
        .Subreport5_usNoIjazahPF.SetUnboundFieldSource ("{ado.NoIjazah}")
        .Subreport5_udTglIjazahPF.SetUnboundFieldSource ("{ado.TglIjazah}")
        .Subreport5_usAlamatPendidikanPF.SetUnboundFieldSource ("{ado.AlamatPendidikan}")
        .Subreport5_usNamaPemimpinPendidikanPF.SetUnboundFieldSource ("{ado.PimpinanPendidikan}")

        'Untuk repot Pendidikan Non Formal
        .Subreport6.OpenSubreport.Database.AddADOCommand dbConn, laporan5
        .Subreport6_usNoUrutNF.SetUnboundFieldSource ("{ado.NoUrut}")
        .Subreport6_usNamaPendidikanNF.SetUnboundFieldSource ("{ado.NamaPendidikan}")
        .Subreport6_usLamaPendidikanNF.SetUnboundFieldSource ("{ado.LamaPendidikan}")
        .Subreport6_udTglMulaiNF.SetUnboundFieldSource ("{ado.TglMulai}")
        .Subreport6_udTglLulusNF.SetUnboundFieldSource ("{ado.TglLulus}")
        .Subreport6_usNoSertifikatNF.SetUnboundFieldSource ("{ado.NoSertifikat}")
        .Subreport6_udTglSertifikatNF.SetUnboundFieldSource ("{ado.TglSertifikat}")
        .Subreport6_usAlamatPendidikanNF.SetUnboundFieldSource ("{ado.AlamatPendidikan}")
        .Subreport6_usKeteranganNF.SetUnboundFieldSource ("{ado.Keterangan}")

        'Untuk repot ExtraPelatihan
        .Subreport7.OpenSubreport.Database.AddADOCommand dbConn, laporan6
        .Subreport7_usNoUrut.SetUnboundFieldSource ("{ado.NoUrut}")
        .Subreport7_usNamaPelatihan.SetUnboundFieldSource ("{ado.NamaPelatihan}")
        .Subreport7_usKedudukanPeranan.SetUnboundFieldSource ("{ado.KedudukanPeranan}")
        .Subreport7_usInstansiPenyelenggara.SetUnboundFieldSource ("{ado.InstansiPenyelenggara}")
        .Subreport7_usAlamatPenyelenggaraan.SetUnboundFieldSource ("{ado.AlamatPenyelenggara}")

        'Untuk repot PengalamanOrganisasi
        .Subreport8.OpenSubreport.Database.AddADOCommand dbConn, laporan9
        .Subreport8_usNoUrut.SetUnboundFieldSource ("{ado.NoUrut}")
        .Subreport8_usNamaOrganisasi.SetUnboundFieldSource ("{ado.NamaOrganisasi}")
        .Subreport8_usJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
        .Subreport8_usTahunAwal.SetUnboundFieldSource ("{ado.TglMasuk}")
        .Subreport8_usTahunAkhir.SetUnboundFieldSource ("{ado.TglAkhir}")
        .Subreport8_usAlamatOrganisasi.SetUnboundFieldSource ("{ado.AlamatOrganisasi}")
        .Subreport8_usNamaPemimpinOrganisasi.SetUnboundFieldSource ("{ado.PimpinanOrganisasi}")
        
        'Untuk repot Pengalamankerja
        .Subreport10.OpenSubreport.Database.AddADOCommand dbConn, laporan10
        .Subreport10_usNoUrut.SetUnboundFieldSource ("{ado.NoUrut}")
        .Subreport10_usNamaPerusahaan.SetUnboundFieldSource ("{ado.NamaPerusahaan}")
        .Subreport10_usJabatanPosisi.SetUnboundFieldSource ("{ado.JabatanPosisi}")
        .Subreport10_usUraianPekerjaan.SetUnboundFieldSource ("{ado.UraianPekerjaan}")
        .Subreport10_udTglMulai.SetUnboundFieldSource ("{ado.TglMulai}")
        .Subreport10_udTglAkhir.SetUnboundFieldSource ("{ado.TglAkhir}")
        .Subreport10_ucGajiPokok.SetUnboundFieldSource ("{ado.GajiPokok}")
        .Subreport10_usNoSuratKeputusan.SetUnboundFieldSource ("{ado.NoSK}")
        .Subreport10_usAlamatPerusahaan.SetUnboundFieldSource ("{ado.AlamatPerusahaan}")

'        .Subreport11.OpenSubreport.Database.AddADOCommand dbConn, laporan11
'        .Subreport11_usNoUrut.SetUnboundFieldSource ("{ado.NoUrut}")
'        .Subreport11_usKotaTujuan.SetUnboundFieldSource ("{ado.KotaTujuan}")
'        .Subreport11_usNegaraTujuan.SetUnboundFieldSource ("{ado.NegaraTujuan}")
'        .Subreport11_usTujuanKunjungan.SetUnboundFieldSource ("{ado.TujuanKunjungan}")
'        .Subreport11_udTglKunjungan.SetUnboundFieldSource ("{ado.TglPergi}")
'        .Subreport11_usLamaKunjungan.SetUnboundFieldSource ("{ado.TglPulang}")
'        .Subreport11_usPenyandangDana.SetUnboundFieldSource ("{ado.PenyandangDana}")
        
        'Untuk repot Prestasi
        .Subreport12.OpenSubreport.Database.AddADOCommand dbConn, laporan12
        .Subreport12_usNoUrut.SetUnboundFieldSource ("{ado.NoUrut}")
        .Subreport12_usNamaPenghargaan.SetUnboundFieldSource ("{ado.NamaPenghargaan}")
        .Subreport12_usTahunDiperoleh.SetUnboundFieldSource ("{ado.TglDiperoleh}")
        .Subreport12_usNamaInstansiPemberi.SetUnboundFieldSource ("{ado.InstansiPemberi}")
        .Subreport12_usKeterangan.SetUnboundFieldSource ("{ado.Keterangan}")

'        .Subreport13.OpenSubreport.Database.AddADOCommand dbConn, laporan13
'        .Subreport13_usKomponenGaji.SetUnboundFieldSource ("{ado.KomponenGaji}")
'        .Subreport13_udTglBerlaku.SetUnboundFieldSource ("{ado.TglSK}")
'        .Subreport13_ucJumlah.SetUnboundFieldSource ("{ado.Jumlah}")

        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
   If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 1
            .DisplayGroupTree = False
        End With
    Else
        Dim tempPrint1 As String
        Dim strDeviceName As String
        Dim strDriverName As String
        Dim strPort As String
        Dim p As Printer
        Dim Posisi, z, Urutan As Integer
        
        Dim sPrinter1 As String
            
            sPrinter1 = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Medifirst2000", "Printer5")
            
                Urutan = 0
                For z = 1 To Len(sPrinter1)
                    If Mid(sPrinter1, z, 1) = ";" Then
                        Urutan = Urutan + 1
                        Posisi = z
                        ReDim Preserve arrPrinter(Urutan)
                        arrPrinter(Urutan).intUrutan = Urutan
                        arrPrinter(Urutan).intPosisi = Posisi
                        If Urutan = 1 Then
                            arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter1, 1, z - 1)
                        Else
                            arrPrinter(Urutan).strNamaPrinter = Mid(sPrinter1, arrPrinter(Urutan - 1).intPosisi + 1, z - arrPrinter(Urutan - 1).intPosisi - 1)
                        End If
                     
                     
                    For Each p In Printers
                            strDeviceName = arrPrinter(Urutan).strNamaPrinter
                            strDriverName = p.DriverName
                            strPort = p.Port
                
                            Report.SelectPrinter strDriverName, strDeviceName, strPort
                            Report.PrintOut False
                            Screen.MousePointer = vbDefault
        
                    Exit For
                    
                    Next
                End If
            Next z
              Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakBukuCVPegawai = Nothing
End Sub
