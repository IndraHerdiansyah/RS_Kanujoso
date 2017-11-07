VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSuratKeteranganCutiTahunan 
   Caption         =   "Cetak Surat Keterangan Cuti"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakSuratKeteranganCutiTahunan.frx":0000
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
      EnableGroupTree =   0   'False
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
Attribute VB_Name = "frmCetakSuratKeteranganCutiTahunan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ReportCutiTahunan As New crSuratKeteranganCutiTahunan
Dim ReportCutiBesar As New crSuratKeteranganCutiBesar
Dim ReportCutiBersalin As New crSuratKeteranganCutiBersalin
Dim ReportCutiPenting As New crSuratKeteranganCutiPenting
Dim ReportCutiSakit As New crSuratKeteranganCutiSakit

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    On Error GoTo errLoad

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    
    ReportCutiTahunan.txtNamaKota.SetText strNKotaRS
    ReportCutiBesar.txtNamaKota.SetText strNKotaRS
    ReportCutiBersalin.txtNamaKota.SetText strNKotaRS
    ReportCutiPenting.txtNamaKota.SetText strNKotaRS
    ReportCutiSakit.txtNamaKota.SetText strNKotaRS

    If frmRiwayatPegawai.dgRiwayatSta.Columns("KdStatus").Value = "02" Then
        With ReportCutiTahunan

            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            .Database.AddADOCommand dbConn, adocomd

            .txtNama.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNama2.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtNIP2.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtJabatan.SetText IIf(IsNull(rs.Fields("NamaJabatan").Value), "", rs.Fields("NamaJabatan").Value)
            .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat Golongan").Value), "", rs.Fields("Pangkat Golongan").Value)
'            .txtKeterangan.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
            .txtKeterangan.SetText IIf(IsNull(rs.Fields("AlamatLengkap").Value), "", rs.Fields("AlamatLengkap").Value)
            .txtTglAwal.SetText IIf(IsNull(Format(rs.Fields("Tglawal").Value, "dd MMMM yyyy")), "", Format(rs.Fields("Tglawal").Value, "dd MMMM yyyy"))
            .txtTglAkhir.SetText IIf(IsNull(Format(rs.Fields("Tglakhir").Value, "dd MMMM yyyy")), "", Format(rs.Fields("Tglakhir").Value, "dd MMMM yyyy"))
            .txtAngkaHari.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", rs.Fields("Lamanya").Value)
            .txtHurufHari.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", NumToText(rs.Fields("Lamanya").Value))

        End With

        CRViewer1.ReportSource = ReportCutiTahunan
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
    End If

    If frmRiwayatPegawai.dgRiwayatSta.Columns("KdStatus").Value = "07" Then
        With ReportCutiBesar

            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            .Database.AddADOCommand dbConn, adocomd

            .txtNama.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNama2.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtNIP2.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtJabatan.SetText IIf(IsNull(rs.Fields("NamaJabatan").Value), "", rs.Fields("NamaJabatan").Value)
            .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat Golongan").Value), "", rs.Fields("Pangkat Golongan").Value)
            .txtKeterangan.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
            .txtTglAwal.SetText IIf(IsNull(Format(rs.Fields("Tglawal").Value, "dd MMMM yyyy")), "", Format(rs.Fields("Tglawal").Value, "dd MMMM yyyy"))
            .txtTglAkhir.SetText IIf(IsNull(Format(rs.Fields("Tglakhir").Value, "dd MMMM yyyy")), "", Format(rs.Fields("Tglakhir").Value, "dd MMMM yyyy"))

        End With

        CRViewer1.ReportSource = ReportCutiBesar
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
    End If

    If frmRiwayatPegawai.dgRiwayatSta.Columns("KdStatus").Value = "09" Then
        With ReportCutiPenting
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            .Database.AddADOCommand dbConn, adocomd

            .txtNama.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtNama2.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP2.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtJabatan.SetText IIf(IsNull(rs.Fields("NamaJabatan").Value), "", rs.Fields("NamaJabatan").Value)
            .txtalasan.SetText IIf(IsNull(rs.Fields("AlasanKeperluan").Value), "", rs.Fields("AlasanKeperluan").Value)
            .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat Golongan").Value), "", rs.Fields("Pangkat Golongan").Value)
            .txtKeterangan.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
            .txtTglAwal.SetText IIf(IsNull(Format(rs.Fields("Tglawal").Value, "dd MMMM yyyy")), "", Format(rs.Fields("Tglawal").Value, "dd MMMM yyyy"))
            .txtTglAkhir.SetText IIf(IsNull(Format(rs.Fields("Tglakhir").Value, "dd MMMM yyyy")), "", Format(rs.Fields("Tglakhir").Value, "dd MMMM yyyy"))
            .txtAngkaHari.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", rs.Fields("Lamanya").Value)
            .txtHurufHari.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", NumToText(rs.Fields("Lamanya").Value))

        End With

        CRViewer1.ReportSource = ReportCutiPenting
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
    End If

    If frmRiwayatPegawai.dgRiwayatSta.Columns("KdStatus").Value = "08" Then
        With ReportCutiBersalin

            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            .Database.AddADOCommand dbConn, adocomd

            .txtNama.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtNama2.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP2.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtJabatan.SetText IIf(IsNull(rs.Fields("NamaJabatan").Value), "", rs.Fields("NamaJabatan").Value)
            .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat Golongan").Value), "", rs.Fields("Pangkat Golongan").Value)
            .txtTglAwal.SetText IIf(IsNull(Format(rs.Fields("Tglawal").Value, "dd MMMM yyyy")), "", Format(rs.Fields("Tglawal").Value, "dd MMMM yyyy"))
            .txtTglAkhir.SetText IIf(IsNull(Format(rs.Fields("Tglakhir").Value, "dd MMMM yyyy")), "", Format(rs.Fields("Tglakhir").Value, "dd MMMM yyyy"))

        End With

        CRViewer1.ReportSource = ReportCutiBersalin
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
    End If

    If frmRiwayatPegawai.dgRiwayatSta.Columns("KdStatus").Value = "10" Then
        With ReportCutiSakit

            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            .Database.AddADOCommand dbConn, adocomd

            .txtNama.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtNama2.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP2.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtJabatan.SetText IIf(IsNull(rs.Fields("NamaJabatan").Value), "", rs.Fields("NamaJabatan").Value)
            .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat Golongan").Value), "", rs.Fields("Pangkat Golongan").Value)

        End With

        CRViewer1.ReportSource = ReportCutiSakit
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
    End If

    If frmRiwayatPegawai.dgRiwayatSta.Columns("KdStatus").Value = "01" Then
        With ReportCutiSakit

            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            .Database.AddADOCommand dbConn, adocomd

            .txtNama.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtNama2.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
            .txtNIP2.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
            .txtJabatan.SetText IIf(IsNull(rs.Fields("NamaJabatan").Value), "", rs.Fields("NamaJabatan").Value)
            .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat Golongan").Value), "", rs.Fields("Pangkat Golongan").Value)

        End With

        CRViewer1.ReportSource = ReportCutiSakit
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
    End If

    Exit Sub

errLoad:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakSuratKeteranganCutiTahunan = Nothing
End Sub
