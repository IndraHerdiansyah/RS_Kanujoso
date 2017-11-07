VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_LaporanStatusPegawai 
   Caption         =   "Cetak Laporan Hukuman Disiplin Pegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm_cetak_LaporanStatusPegawai.frx":0000
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
Attribute VB_Name = "frm_cetak_LaporanStatusPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crLaporanPegawaiCuti


Private Sub Form_Load()
On Error GoTo hell

Screen.MousePointer = vbHourglass
Me.WindowState = 2
Set frm_cetak_LaporanStatusPegawai = Nothing
Dim adocomd As New ADODB.Command
Call openConnection

adocomd.ActiveConnection = dbConn

adocomd.CommandText = strSQL

adocomd.CommandType = adCmdText
Report.Database.AddADOCommand dbConn, adocomd
Report.txtNamaRS.SetText strNNamaRS
Report.txtAlamat.SetText strNAlamatRS
'Report.Text3.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
Report.txtTanggal1.SetText Format(frmLaporanStatusCuti.DTPickerAwal.Value, "MMMM yyyy")
Report.txtTanggal2.SetText Format(frmLaporanStatusCuti.DTPickerAkhir.Value, "MMMM yyyy")


Report.usNamaPegawai.SetUnboundFieldSource ("{ado.Nama}")
Report.usNIP.SetUnboundFieldSource ("{ado.NIP}")
Report.usJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
Report.usRuangan.SetUnboundFieldSource ("{ado.Ruangan Tugas}")
Report.usStatus.SetUnboundFieldSource ("{ado.Status}")
Report.udMulai.SetUnboundFieldSource ("{ado.Tgl Mulai}")
Report.udAkhir.SetUnboundFieldSource ("{ado.Tgl Akhir}")
Report.usAlasan.SetUnboundFieldSource ("{ado.Alasan Keperluan}")

'strSQLsplakuk = "select NamaLengkap, NamaPangkat, NIP from V_FooterPegawai where KdJabatan='A02'"
'Call msubRecFO(rsSplakuk, strSQLsplakuk)
'
'Report.txtDirektur.SetText IIf(IsNull(rsSplakuk("NamaLengkap").Value), "", rsSplakuk("NamaLengkap").Value)
'Report.txtPangkatF.SetText IIf(IsNull(rsSplakuk("NamaPangkat").Value), "", rsSplakuk("NamaPangkat").Value)
'Report.txtNIPF.SetText IIf(IsNull(rsSplakuk("NIP").Value), "", rsSplakuk("NIP").Value)

Screen.MousePointer = vbHourglass
With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
End With
Screen.MousePointer = vbDefault
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_cetak_LaporanStatusPegawai = Nothing
End Sub

