VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_LaporanRiwayatKirimTerimaTenagaPemeriksa 
   Caption         =   "Cetak Riwayat Kirim Terima Tenaga Pemeriksa"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm_cetak_LaporanRiwayatKirimTerimaTenagaPemeriksa.frx":0000
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
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frm_cetak_LaporanRiwayatKirimTerimaTenagaPemeriksa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crLaporanRiwayatKirimTerimaTenagaPemeriksa

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Set frm_cetak_LaporanRiwayatKirimTerimaTenagaPemeriksa = Nothing
    Dim adocomd As New ADODB.Command

    adocomd.ActiveConnection = dbConn

    adocomd.CommandText = strSQL

    adocomd.CommandType = adCmdText

    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS
        If frmRiwayatKirimDanTerimaTenagaPemeriksa.optKirim.Value = True Then
            .Text1.SetText "RIWAYAT KIRIM TENAGA PEMERIKSA"
            .Text9.SetText "Tgl. Kirim"
            .usNamaDokter.SetUnboundFieldSource ("{ado.NamaDokter}")
            .usRujukanTujuan.SetUnboundFieldSource ("{Ado.RujukanTujuan}")
            .usKualifikasi.SetUnboundFieldSource ("{Ado.KualifikasiJurusan}")
            .usSubInstalasi.SetUnboundFieldSource ("{ado.SubInstalasi}")
            .usTempatRujukan.SetUnboundFieldSource ("{ado.TempatRujukanTujuan}")
            .usAlamat.SetUnboundFieldSource ("{ado.AlamatTempatRujukanTujuan}")
            .udtTglKirim.SetUnboundFieldSource ("{Ado.TglKirim}")
            .udtTglKembali.SetUnboundFieldSource ("{ado.TglKembali}")
            .usKeterangan.SetUnboundFieldSource ("{ado.Keterangan}")
        End If
        If frmRiwayatKirimDanTerimaTenagaPemeriksa.optTerima.Value = True Then
            .Text1.SetText "RIWAYAT TERIMA TENAGA PEMERIKSA"
            .Text9.SetText "Tgl. Terima"
            .usNamaDokter.SetUnboundFieldSource ("{ado.NamaDokter}")
            .usRujukanTujuan.SetUnboundFieldSource ("{Ado.RujukanAsal}")
            .usKualifikasi.SetUnboundFieldSource ("{Ado.KualifikasiJurusan}")
            .usSubInstalasi.SetUnboundFieldSource ("{ado.SubInstalasi}")
            .usTempatRujukan.SetUnboundFieldSource ("{ado.TempatRujukanAsal}")
            .usAlamat.SetUnboundFieldSource ("{ado.AlamatTempatRujukanAsal}")
            .udtTglKirim.SetUnboundFieldSource ("{Ado.TglTerima}")
            .udtTglKembali.SetUnboundFieldSource ("{ado.TglKembali}")
            .usKeterangan.SetUnboundFieldSource ("{ado.Keterangan}")
        End If

    End With
    Screen.MousePointer = vbHourglass

    If vLaporan = "Print" Then
        Report.PrintOut False
        Unload Me
    Else

        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 100
        End With
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
    Set frm_cetak_LaporanRiwayatKirimTerimaTenagaPemeriksa = Nothing
End Sub
