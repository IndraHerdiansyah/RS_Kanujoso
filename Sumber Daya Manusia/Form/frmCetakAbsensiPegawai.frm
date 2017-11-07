VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakAbsensiPegawai 
   Caption         =   "frmCetakAbsensiPegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakAbsensiPegawai.frx":0000
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
Attribute VB_Name = "frmCetakAbsensiPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crAbsensiPegawai
Dim ReportDetail As New crAbsensiPegawaiDetail
Dim Judul1, Judul2, Judul3, Judul4 As String
Dim intJmlHari As Integer

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn

    Judul1 = "LAPORAN ABSENSI PEGAWAI (PER HARI)"
    Judul2 = "LAPORAN ABSENSI PEGAWAI (PER BULAN)"
    Judul3 = "LAPORAN ABSENSI PEGAWAI (PER JAM)"
    Judul4 = "LAPORAN ABSENSI PEGAWAI (PER TAHUN)"

    Select Case strCetak
        Case "Hari"
            Call LaporanPerHari
        Case "Bulan"
            Call LaporanPerBulan
        Case "Tahun"
            Call LaporanPerBulan
        Case "Jam"
            Call LaporanPerJam
    End Select
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakAbsensiPegawai = Nothing
End Sub

Private Sub LaporanPerHari()
    On Error GoTo hell
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn

    If strCetak2 = "CetakAbsensi" Then
        'strSQL = "SELECT NamaLengkap, IdPegawai, NIP, NamaJabatan,Total" & _
        " From v_CetakAbsensi2" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND WaktuMasuk BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "
        strSQL = "SELECT NamaLengkap, IdPegawai, NIP, NamaJabatan,Total" & _
        " From v_CetakAbsensi3" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND WaktuMasuk BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "

        Report.txtGroup.SetText strGroup
        Report.txtIsiGroup.SetText strIsiGroup

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        Report.Database.AddADOCommand dbConn, adocomd
        With Report
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtAlamatRS.SetText strNAlamatRS & " " & strNKotaRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy") & ""
            .txtJudul.SetText Judul1
            
            .txtNamaKota.SetText strNKotaRS

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .UsJabatan.SetUnboundFieldSource ("{ado.NamaJabatan}")
            .udtDate.SetUnboundFieldSource ("{ado.TglAbsen}")
            '.udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
            .unTotal.SetUnboundFieldSource ("{ado.Total}")
            If .unTotal.Value > "80" Then .unTotal.BackColor = vbGreen
        End With

        If vLaporan = "Print" Then
            Report.PrintOut False
            Unload Me
        Else
            With CRViewer1
                .ReportSource = Report
                .EnableGroupTree = False
                .EnableExportButton = True
                .ViewReport
                .Zoom 100
            End With
        End If

        '-----------cetak detail

    ElseIf strCetak2 = "CetakDetailAbsensi" Then
        strSQL = "SELECT NamaLengkap, NIP, JenisPegawai, Jabatan, Instalasi, Ruangan, TglAbsen " & _
        " From v_CetakAbsensiDetail" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND TglAbsen BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        ReportDetail.Database.AddADOCommand dbConn, adocomd
        With ReportDetail
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtAlamatRS.SetText strNAlamatRS & " " & strNKotaRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy") & ""
            .txtJudul.SetText Judul1
            
            .txtNamaKota.SetText strNKotaRS

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .UsNIP.SetUnboundFieldSource ("{ado.NIP}")
            .UsJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
            .udtTglMasuk.SetUnboundFieldSource ("{ado.TglAbsen}")
            '.udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")

            If frmLaporanDetailAbsensi.chkGroup.Value = vbChecked Then
                If frmLaporanDetailAbsensi.optInstalasi.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Instalasi}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                    .Text9.Suppress = False: .UnboundNumber2.Suppress = False
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                End If
            ElseIf frmLaporanDetailAbsensi.chkGroup.Value = vbUnchecked Then
                If frmLaporanDetailAbsensi.optInstalasi.Value = True Then
                    ReportDetail.txtGroup.SetText "Instalasi :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Instalasi}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText "Ruangan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Nama :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                    .Text9.Suppress = False: .UnboundNumber2.Suppress = False
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Jenis Pegawai :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText "Jabatan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                End If
            End If
        End With

        If vLaporan = "Print" Then
            ReportDetail.PrintOut False
            Unload Me
        Else
            With CRViewer1
                .ReportSource = ReportDetail
                .EnableGroupTree = False
                .EnableExportButton = True
                .ViewReport
                .Zoom 100
            End With
        End If
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub LaporanPerBulan()
    On Error GoTo hell
    Dim adocomd As New ADODB.Command

    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn

    If strCetak2 = "CetakAbsensi" Then
        If (strCetak = "Tahun") Then
            'strSQL = "SELECT NamaInstalasi,NamaRuangan, NamaLengkap, IdPegawai, NIP, NamaJabatan, {fn MONTHNAME (WaktuMasuk)} As TglAbsen,convert(char(4),year(WaktuMasuk)) as Tahun,WaktuMasuk,WaktuKeluar,Total" & _
            " From v_CetakAbsensi2" & _
            " WHERE " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND YEAR(WaktuMasuk) between '" & (Format(mdTglAwal, "yyyy")) & "' AND '" & (Format(mdTglAkhir, "yyyy")) & "'"
            strSQL = "SELECT NamaInstalasi,NamaRuangan, NamaLengkap, IdPegawai, NIP, NamaJabatan, {fn MONTHNAME (WaktuMasuk)} As TglAbsen,convert(char(4),year(WaktuMasuk)) as Tahun,WaktuMasuk,WaktuKeluar,Total" & _
            " From v_CetakAbsensi3" & _
            " WHERE " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND YEAR(WaktuMasuk) between '" & (Format(mdTglAwal, "yyyy")) & "' AND '" & (Format(mdTglAkhir, "yyyy")) & "'"
        Else
            'strSQL = "SELECT NamaInstalasi,NamaRuangan, NamaLengkap, IdPegawai, NIP, NamaJabatan, {fn MONTHNAME (WaktuMasuk)} As TglAbsen,YEAR(WaktuMasuk) as Tahun,WaktuMasuk,WaktuKeluar,Total" & _
            " From v_CetakAbsensi2" & _
            " WHERE " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND MONTH(WaktuMasuk) between '" & Month(Format(mdTglAwal, "dd/MM/yyyy")) & "' AND '" & Month(Format(mdTglAkhir, "dd/MM/yyyy")) & "'" & _
            " AND YEAR(WaktuMasuk) ='" & Year(mdTglAkhir) & "'"
            strSQL = "SELECT NamaInstalasi,NamaRuangan, NamaLengkap, IdPegawai, NIP, NamaJabatan, {fn MONTHNAME (WaktuMasuk)} As TglAbsen,YEAR(WaktuMasuk) as Tahun,WaktuMasuk,WaktuKeluar,Total" & _
            " From v_CetakAbsensi3" & _
            " WHERE " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND MONTH(WaktuMasuk) between '" & Month(Format(mdTglAwal, "dd/MM/yyyy")) & "' AND '" & Month(Format(mdTglAkhir, "dd/MM/yyyy")) & "'" & _
            " AND YEAR(WaktuMasuk) ='" & Year(mdTglAkhir) & "'"
        End If
        


        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        Report.Database.AddADOCommand dbConn, adocomd

        With Report
            
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtAlamatRS.SetText strNAlamatRS & " " & strNKotaRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            If (strCetak = "Tahun") Then
                .txtPeriode.SetText "Tahun : " & Format(mdTglAwal, "yyyy") & " - " & Format(mdTglAkhir, "yyyy")
                .txtJudul.SetText Judul4
            Else
                .txtPeriode.SetText "Bulan : " & Format(mdTglAwal, "MMMM yyyy") & " - " & Format(mdTglAkhir, "MMMM yyyy")
                .txtJudul.SetText Judul2
            End If
                         
            .txtNamaKota.SetText strNKotaRS
            If (strGroup = "NamaRuangan") Then
                .usGroup.SetUnboundFieldSource ("{ado.NamaRuangan}")
            ElseIf (strGroup = "NamaInstalasi") Then
                .usGroup.SetUnboundFieldSource ("{ado.NamaInstalasi}")
            End If
            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .UsJabatan.SetUnboundFieldSource ("{ado.NamaJabatan}")
            If (strCetak = "Bulan") Then
                .udtDate.SetUnboundFieldSource ("{ado.TglAbsen}")
            Else
                .udtDate.SetUnboundFieldSource ("{ado.Tahun}")
            End If
            '.udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
            .unTotal.SetUnboundFieldSource ("{ado.Total}")
        End With

        If vLaporan = "Print" Then
            Report.PrintOut False
            Unload Me
        Else
            With CRViewer1
                .ReportSource = Report
                .EnableGroupTree = False
                .EnableExportButton = True
                .ViewReport
                .Zoom 100
            End With
        End If

    ElseIf strCetak2 = "CetakDetailAbsensi" Then
        If (strCetak = "Tahun") Then
            'strSQL = "SELECT convert(varchar(3),Total)+' Jam '+ convert(varchar(2),TotalMenit)+' Menit' as Total,IdPegawai,NamaLengkap, NIP,NamaJabatan, NamaInstalasi, NamaRuangan, WaktuMasuk,WaktuKeluar  " &
            '//yayang.agus 2014-08-11
            'strSQL = "SELECT dbo.selisih_tanggal(waktumasuk,waktukeluar) as Total,IdPegawai,NamaLengkap, NIP,NamaJabatan, NamaInstalasi, NamaRuangan, WaktuMasuk,WaktuKeluar  " & _
            " From v_CetakAbsensi2 " & _
            " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND Year(WaktuMasuk) BETWEEN '" & Year(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Year(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "' order by NamaLengkap"
'            strSQL = "SELECT dbo.selisih_tanggal(waktumasuk,waktukeluar) as Total,IdPegawai,NamaLengkap, NIP,NamaJabatan, NamaInstalasi, NamaRuangan, WaktuMasuk,WaktuKeluar  " & _
            " From v_CetakAbsensi3 " & _
            " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND Year(WaktuMasuk) BETWEEN '" & Year(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Year(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "' order by NamaLengkap"
            strSQL = "SELECT dbo.JmlJamKerja(waktumasuk,waktukeluar,jamistirahatawal,jamistirahatakhir) as Total,IdPegawai,NamaLengkap, NIP,NamaJabatan, NamaInstalasi, NamaRuangan, WaktuMasuk,WaktuKeluar  " & _
            " From v_CetakAbsensi3 " & _
            " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND Year(WaktuMasuk) BETWEEN '" & Year(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Year(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "' order by NamaLengkap"
        Else
            'strSQL = "SELECT convert(varchar(3),Total)+' Jam '+ convert(varchar(2),TotalMenit)+' Menit' as Total,IdPegawai,NamaLengkap, NIP,NamaJabatan, NamaInstalasi, NamaRuangan, WaktuMasuk,WaktuKeluar  " &
            '//yayang.agus 2014-08-11
            'strSQL = "SELECT dbo.selisih_tanggal(waktumasuk,waktukeluar) as Total,IdPegawai,NamaLengkap, NIP,NamaJabatan, NamaInstalasi, NamaRuangan, WaktuMasuk,WaktuKeluar  " & _
            " From v_CetakAbsensi2 " & _
            " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND MONTH(WaktuMasuk) BETWEEN '" & Month(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Month(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "'" & _
            " AND YEAR(WaktuMasuk) BETWEEN '" & Year(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Year(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "' order by NamaLengkap"
'            strSQL = "SELECT dbo.selisih_tanggal(waktumasuk,waktukeluar) as Total,IdPegawai,NamaLengkap, NIP,NamaJabatan, NamaInstalasi, NamaRuangan, WaktuMasuk,WaktuKeluar  " & _
            " From v_CetakAbsensi3 " & _
            " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND MONTH(WaktuMasuk) BETWEEN '" & Month(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Month(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "'" & _
            " AND YEAR(WaktuMasuk) BETWEEN '" & Year(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Year(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "' order by NamaLengkap"
            strSQL = "SELECT dbo.JmlJamKerja(waktumasuk,waktukeluar,jamistirahatawal,jamistirahatakhir) as Total,IdPegawai,NamaLengkap, NIP,NamaJabatan, NamaInstalasi, NamaRuangan, WaktuMasuk,WaktuKeluar  " & _
            " From v_CetakAbsensi3 " & _
            " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
            " AND MONTH(WaktuMasuk) BETWEEN '" & Month(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Month(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "'" & _
            " AND YEAR(WaktuMasuk) BETWEEN '" & Year(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Year(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "' order by NamaLengkap"
        End If
        

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        ReportDetail.Database.AddADOCommand dbConn, adocomd
        With ReportDetail

            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtAlamatRS.SetText strNAlamatRS & " " & strNKotaRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            If (strCetak = "Tahun") Then
                .txtPeriode.SetText "Tahun : " & Format(mdTglAwal, "yyyy") & " - " & Format(mdTglAkhir, "yyyy")
                .txtJudul.SetText Judul4
            Else
                .txtPeriode.SetText "Bulan : " & Format(mdTglAwal, "MMMM yyyy") & " - " & Format(mdTglAkhir, "MMMM yyyy")
                .txtJudul.SetText Judul2
            End If
'            .txtPeriode.SetText "Bulan : " & Format(mdTglAwal, "MMMM yyyy") & " - " & Format(mdTglAkhir, "MMMM yyyy")
'            .txtJudul.SetText Judul2
            
            .txtNamaKota.SetText strNKotaRS

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .UsNIP.SetUnboundFieldSource ("{ado.NIP}")
            .UsJabatan.SetUnboundFieldSource ("{ado.NamaJabatan}")
            .udtTglMasuk.SetUnboundFieldSource ("{ado.WaktuMasuk}")
            .udtTglPulang.SetUnboundFieldSource ("{ado.WaktuKeluar}")
            .usTotalJam.SetUnboundFieldSource ("{ado.Total}")
            If frmLaporanDetailAbsensi.chkGroup.Value = vbChecked Then
                If frmLaporanDetailAbsensi.optInstalasi.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaInstalasi}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaRuangan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                    .Text9.Suppress = False: .UnboundNumber2.Suppress = False
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaJabatan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                End If
            ElseIf frmLaporanDetailAbsensi.chkGroup.Value = vbUnchecked Then
                If frmLaporanDetailAbsensi.optInstalasi.Value = True Then
                    ReportDetail.txtGroup.SetText "NamaInstalasi :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaInstalasi}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText "NamaRuangan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaRuangan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Nama :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                    .Text9.Suppress = False: .UnboundNumber2.Suppress = False
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Jenis Pegawai :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText "Jabatan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaJabatan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                End If
            End If
        End With

        If vLaporan = "Print" Then
            ReportDetail.PrintOut False
            Unload Me
        Else
            With CRViewer1
                .ReportSource = ReportDetail
                .EnableGroupTree = False
                .EnableExportButton = True
                .ViewReport
                .Zoom 100
            End With
        End If
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub LaporanPerJam()
    On Error GoTo hell
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn

    If strCetak2 = "CetakAbsensi" Then

        'strSQL = "SELECT NamaLengkap, NamaJabatan, TglAbsen, Total" & _
        " From v_CetakAbsensi2" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND WaktuMasuk BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd HH:mm:ss") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd HH:mm:ss") & "' "
        strSQL = "SELECT NamaLengkap, NamaJabatan, TglAbsen, Total" & _
        " From v_CetakAbsensi3" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND WaktuMasuk BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd HH:mm:ss") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd HH:mm:ss") & "' "
        Report.txtGroup.SetText strGroup
        Report.txtIsiGroup.SetText strIsiGroup

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        Report.Database.AddADOCommand dbConn, adocomd
        With Report
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtAlamatRS.SetText strNAlamatRS & " " & strNKotaRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy HH:mm:ss") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm:ss") & ""
            .txtJudul.SetText Judul3
            
            .txtNamaKota.SetText strNKotaRS

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .UsJabatan.SetUnboundFieldSource ("{ado.NamaJabatan}")
            '.udtTglMasuk.SetUnboundFieldSource ("{ado.TglAbsen}")
            '.udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
            '.unTotal.SetUnboundFieldSource ("{ado.Total}")
        End With

        If vLaporan = "Print" Then
            Report.PrintOut False
            Unload Me
        Else
            With CRViewer1
                .ReportSource = Report
                .EnableGroupTree = False
                .EnableExportButton = True
                .ViewReport
                .Zoom 100
            End With
        End If

    ElseIf strCetak2 = "CetakDetailAbsensi" Then
        strSQL = "SELECT NamaLengkap, NIP, JenisPegawai, Jabatan, Instalasi, Ruangan, TglAbsen" & _
        " From v_CetakAbsensiDetail" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND TglAbsen BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd HH:mm") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd HH:mm") & "' "

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        ReportDetail.Database.AddADOCommand dbConn, adocomd
        With ReportDetail
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtAlamatRS.SetText strNAlamatRS & " " & strNKotaRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy HH:mm:ss") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm:ss") & ""
            .txtJudul.SetText Judul3
            
            .txtNamaKota.SetText strNKotaRS

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .UsJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
            .udtTglMasuk.SetUnboundFieldSource ("{ado.TglAbsen}")
            '.udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")

            If frmLaporanDetailAbsensi.chkGroup.Value = vbChecked Then
                If frmLaporanDetailAbsensi.optInstalasi.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Instalasi}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                    .Text9.Suppress = False: .UnboundNumber2.Suppress = False
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                End If
            ElseIf frmLaporanDetailAbsensi.chkGroup.Value = vbUnchecked Then
                If frmLaporanDetailAbsensi.optInstalasi.Value = True Then
                    ReportDetail.txtGroup.SetText "Instlasi :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Instalasi}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText "Ruangan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Nama :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                    .Text9.Suppress = False: .UnboundNumber2.Suppress = False
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Jenis Pegawai :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText "Jabatan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                    .Text9.Suppress = True: .UnboundNumber2.Suppress = True
                End If
            End If
        End With

        If vLaporan = "Print" Then
            ReportDetail.PrintOut False
            Unload Me
        Else
            With CRViewer1
                .ReportSource = ReportDetail
                .EnableGroupTree = False
                .EnableExportButton = True
                .ViewReport
                .Zoom 100
            End With
        End If
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub
