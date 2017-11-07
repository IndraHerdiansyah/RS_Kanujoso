VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakStatusAbsenPegawai 
   Caption         =   "frmCetakAbsensiPegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakStatusAbsenPegawai.frx":0000
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
Attribute VB_Name = "frmCetakStatusAbsenPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crAbsensiPegawai
Dim ReportDetail As New crAbsensiPegawaiDetail
Dim Judul1, Judul2, Judul3 As String
Dim intJmlHari As Integer

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn

    Judul1 = "LAPORAN ABSENSI PEGAWAI (PER HARI)"
    Judul2 = "LAPORAN ABSENSI PEGAWAI (PER BULAN)"
    Judul3 = "LAPORAN ABSENSI PEGAWAI (PER JAM)"

    Select Case strCetak
        Case "Hari"
            Call LaporanPerHari
        Case "Bulan"
            Call LaporanPerBulan
        Case "Jam"
            Call LaporanPerJam
    End Select
    Me.WindowState = 2
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
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn

    If strCetak2 = "CetakAbsensi" Then
        strSQL = "SELECT NamaLengkap, Jabatan, TglMasuk, TglPulang, Total" & _
        " From v_CetakAbsensi" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND TglMasuk BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "

        Report.txtGroup.SetText strGroup
        Report.txtIsiGroup.SetText strIsiGroup

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        Report.Database.AddADOCommand dbConn, adocomd
        With Report
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
            .txtAlamatRS.SetText strNAlamatRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy") & ""
            .txtJudul.SetText Judul1

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .usJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
            .udtTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
            .udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
            .unTotal.SetUnboundFieldSource ("{ado.Total}")
            If .unTotal.Value > "80" Then .unTotal.BackColor = vbGreen
        End With
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .EnableExportButton = True
            .ViewReport
            .Zoom 100
        End With

    ElseIf strCetak2 = "CetakDetailAbsensi" Then
        strSQL = "SELECT NamaLengkap, JenisPegawai, Jabatan, Ruangan, Instalasi, Total, TglMasuk, TglPulang" & _
        " From v_CetakAbsensi" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND TglMasuk BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "' "

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        ReportDetail.Database.AddADOCommand dbConn, adocomd
        With ReportDetail
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
            .txtAlamatRS.SetText strNAlamatRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy") & ""
            .txtJudul.SetText Judul1

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .usJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
            .udtTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
            .udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")

            If frmLaporanDetailAbsensi.chkGroup.Value = vbChecked Then
                If frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                End If
            ElseIf frmLaporanDetailAbsensi.chkGroup.Value = vbUnchecked Then
                If frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText "Ruangan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Nama :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Jenis Pegawai :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText "Jabatan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                End If
            End If
        End With
        With CRViewer1
            .ReportSource = ReportDetail
            .EnableGroupTree = False
            .EnableExportButton = True
            .ViewReport
            .Zoom 100
        End With
    End If
End Sub

Private Sub LaporanPerBulan()
    Dim adocomd As New ADODB.Command

    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn

    If strCetak2 = "CetakAbsensi" Then
        strSQL = "SELECT XXX, NamaLengkap, NIP, Jabatan, {fn MONTHNAME (TglRiwayat)} As TglRiwayat, TglMasuk, TglPulang, Total" & _
        " From v_CetakAbsensi" & _
        " WHERE " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND MONTH(TglMasuk) between '" & Month(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Month(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "'" & _
        " AND YEAR(TglMasuk) ='" & Year(mdTglAkhir) & "'"

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        Report.Database.AddADOCommand dbConn, adocomd

        With Report
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
            .txtAlamatRS.SetText strNAlamatRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Bulan : " & Format(mdTglAwal, "MMMM yyyy") & ""
            .txtJudul.SetText Judul2

            .usNamaPegawai.SetUnboundFieldSource ("{ado.XXX}")
            .usJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
            .udtTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
            .udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
            .unTotal.SetUnboundFieldSource ("{ado.Total}")
        End With
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .EnableExportButton = True
            .ViewReport
            .Zoom 100
        End With

    ElseIf strCetak2 = "CetakDetailAbsensi" Then
        strSQL = "SELECT NamaLengkap, JenisPegawai, Jabatan, Ruangan, Instalasi, {fn MONTHNAME (TglRiwayat)} As TglRiwayat, TglMasuk, TglPulang, Total" & _
        " From v_CetakAbsensi" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND MONTH(TglMasuk) BETWEEN '" & Month(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Month(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "'" & _
        " AND YEAR(TglMasuk) BETWEEN '" & Year(Format(mdTglAwal, "yyyy/MM/01 00:00:00")) & "' AND '" & Year(Format(mdTglAkhir, "yyyy/MM/dd 23:59:59")) & "' order by NamaLengkap"

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        ReportDetail.Database.AddADOCommand dbConn, adocomd
        With ReportDetail
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
            .txtAlamatRS.SetText strNAlamatRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Bulan : " & Format(mdTglAwal, "MMMM yyyy") & ""
            .txtJudul.SetText Judul2

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .usJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
            .udtTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
            .udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
            If frmLaporanDetailAbsensi.chkGroup.Value = vbChecked Then
                If frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                End If
            ElseIf frmLaporanDetailAbsensi.chkGroup.Value = vbUnchecked Then
                If frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText "Ruangan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Nama :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Jenis Pegawai :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText "Jabatan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                End If
            End If
        End With
        With CRViewer1
            .ReportSource = ReportDetail
            .EnableGroupTree = False
            .EnableExportButton = True
            .ViewReport
            .Zoom 100
        End With
    End If
End Sub

Private Sub LaporanPerJam()
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn

    If strCetak2 = "CetakAbsensi" Then
        strSQL = "SELECT NamaLengkap, Jabatan, TglMasuk, TglPulang, Total" & _
        " From v_CetakAbsensi" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND TglMasuk BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd HH:mm:ss") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd HH:mm:ss") & "' "

        Report.txtGroup.SetText strGroup
        Report.txtIsiGroup.SetText strIsiGroup

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        Report.Database.AddADOCommand dbConn, adocomd
        With Report
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
            .txtAlamatRS.SetText strNAlamatRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy HH:mm:ss") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm:ss") & ""
            .txtJudul.SetText Judul3

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .usJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
            .udtTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
            .udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
            .unTotal.SetUnboundFieldSource ("{ado.Total}")

        End With
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .EnableExportButton = True
            .ViewReport
            .Zoom 100
        End With

    ElseIf strCetak2 = "CetakDetailAbsensi" Then
        strSQL = "SELECT NamaLengkap, JenisPegawai, Jabatan, Ruangan, Instalasi, Total, TglMasuk, TglPulang" & _
        " From v_CetakAbsensi" & _
        " Where " & strGroup & " LIKE '%" & strIsiGroup & "%' " & _
        " AND TglMasuk BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd HH:mm:ss") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd HH:mm:ss") & "' "

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        ReportDetail.Database.AddADOCommand dbConn, adocomd
        With ReportDetail
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
            .txtAlamatRS.SetText strNAlamatRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos " & " " & strNKodepos & " "
            .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy HH:mm") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy HH:mm") & ""
            .txtJudul.SetText Judul3

            .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaLengkap}")
            .usJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
            .udtTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
            .udtTglPulang.SetUnboundFieldSource ("{ado.TglPulang}")
            If frmLaporanDetailAbsensi.chkGroup.Value = vbChecked Then
                If frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText strGroup
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                End If
            ElseIf frmLaporanDetailAbsensi.chkGroup.Value = vbUnchecked Then
                If frmLaporanDetailAbsensi.optRuangan.Value = True Then
                    ReportDetail.txtGroup.SetText "Ruangan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Ruangan}")
                ElseIf frmLaporanDetailAbsensi.optPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Nama :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.NamaLengkap}")
                ElseIf frmLaporanDetailAbsensi.optJenisPegawai.Value = True Then
                    ReportDetail.txtGroup.SetText "Jenis Pegawai :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.JenisPegawai}")
                ElseIf frmLaporanDetailAbsensi.optJabatan.Value = True Then
                    ReportDetail.txtGroup.SetText "Jabatan :"
                    .usIsiGroup.SetUnboundFieldSource ("{ado.Jabatan}")
                End If
            End If
        End With
        With CRViewer1
            .ReportSource = ReportDetail
            .EnableGroupTree = False
            .EnableExportButton = True
            .ViewReport
            .Zoom 100
        End With
    End If
End Sub

