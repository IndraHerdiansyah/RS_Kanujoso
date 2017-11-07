VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakRekapLaporanGaji 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Cetak Rekap Laporan Gaji"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19155
   Icon            =   "frmCetakRekapLaporanGaji.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   19155
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19125
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
Attribute VB_Name = "frmCetakRekapLaporanGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As crRekapLaporanGajiPegawai
Dim SubReportGP As crSubRekapLaporanGajiPegawaiGP

Private Sub Form_Load()
    If frmRekapLaporanGajiNew.chkPerbagian.Value = 1 Then
        'Call perbagian
    Else
        Call rekapAll
    End If
    Screen.MousePointer = vbHourglass
    Me.WindowState = 0
    Screen.MousePointer = vbDefault
End Sub

Public Sub rekapAll()
    Dim totalpenerimaan As Currency
    Dim totalPotongan As Currency
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    Set Report = New crRekapLaporanGajiPegawai

    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    End With
    strSQL = "Select tglPembayaran,idPegawai,[Nama Lengkap], KomponenGaji, Jumlah, komponenpotongangaji, jumlahpotongan from V_PembayaranGajiDanPotonganPegawai where tglPembayaran between '" & Format(frmRekapLaporanGajiNew.dtpTglAwal.Value, "yyyy/MM/dd") & "'and '" & Format(frmRekapLaporanGajiNew.dtpTglAhir.Value, "yyyy/MM/dd") & "' ORDER BY tglPembayaran"
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, dbcmd
    Report.usTglPembayaran.SetUnboundFieldSource ("{ado.tglPembayaran}")
    Report.usNamaPegawai.SetUnboundFieldSource ("{ado.Nama Lengkap}")
    Report.usNamaPenerimaan.SetUnboundFieldSource ("{ado.KomponenGaji}")
    Report.usjmlPenerimaan.SetUnboundFieldSource ("{ado.Jumlah}")
    Report.usNamaPotongan.SetUnboundFieldSource ("{ado.KomponenPotonganGaji}")
    Report.usjmlPotongan.SetUnboundFieldSource ("{ado.JumlahPotongan}")
    Set rs = Nothing
    strSQL = "SELECT SUM(Jumlah) As JumlahPenerimaan from V_PembayaranGajiDanPotonganPegawai where tglPembayaran between '" & Format(frmRekapLaporanGajiNew.dtpTglAwal.Value, "yyyy/MM/dd") & "'and '" & Format(frmRekapLaporanGajiNew.dtpTglAhir.Value, "yyyy/MM/dd") & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    totalpenerimaan = IIf(IsNull(rs(0)), 0, rs(0))
    Report.totalpenerimaan.SetText totalpenerimaan
    Set rs = Nothing
    strSQL = "SELECT SUM(JumlahPotongan) As JumlahPotongan from V_PembayaranGajiDanPotonganPegawai where tglPembayaran between '" & Format(frmRekapLaporanGajiNew.dtpTglAwal.Value, "yyyy/MM/dd") & "'and '" & Format(frmRekapLaporanGajiNew.dtpTglAhir.Value, "yyyy/MM/dd") & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    totalPotongan = IIf(IsNull(rs(0)), 0, rs(0))
    Report.totalPotongan.SetText totalPotongan
    Set rs = Nothing
    strSQL = "SELECT SUM(Jumlah) As JumlahPenerimaan from V_PembayaranGajiDanPotonganPegawai where tglPembayaran between '" & Format(frmRekapLaporanGajiNew.dtpTglAwal.Value, "yyyy/MM/dd") & "'and '" & Format(frmRekapLaporanGajiNew.dtpTglAhir.Value, "yyyy/MM/dd") & "' and kdKomponenGaji='01'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Report.TotalGP.SetText IIf(IsNull(rs(0)), 0, rs(0))
    Set rs = Nothing
    strSQL = "SELECT SUM(JumlahPotongan) As JumlahPPH from V_PembayaranGajiDanPotonganPegawai where tglPembayaran between '" & Format(frmRekapLaporanGajiNew.dtpTglAwal.Value, "yyyy/MM/dd") & "'and '" & Format(frmRekapLaporanGajiNew.dtpTglAhir.Value, "yyyy/MM/dd") & "' and KomponenPotonganGaji like 'PPH%'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Report.TotalPPH.SetText IIf(IsNull(rs(0)), 0, rs(0))

    CRViewer1.ReportSource = Report
    With CRViewer1
        .EnableGroupTree = False
        .ViewReport
        .Zoom 1
    End With
End Sub

Private Sub Form_Resize()
    With CRViewer1
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakRekapLaporanGaji = Nothing
End Sub
