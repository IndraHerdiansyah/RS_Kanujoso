VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakIndexPegawaiPerRuangan 
   Caption         =   "Form Cetak Index Pegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakIndexPegawaiPerRuangan.frx":0000
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
Attribute VB_Name = "frmCetakIndexPegawaiPerRuangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crIndexPegawaiPerRuangan

Private Sub Form_Load()
On Error GoTo errLoad
Dim i As Integer
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn
    strSQL = "Select * from vJasaRemunerasiPegawai WHERE NamaKomputer = '" & strNamaHostLocal & "'  AND Bulan= '" & MonthName(Month(frmInsentifKaryawan.dtpPeriode.Value)) & "' and Tahun='" & Year(frmInsentifKaryawan.dtpPeriode.Value) & "'"
    Call msubRecFO(rs, strSQL)
    
    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText

    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
        .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
        .txtAlamatRS.SetText strNAlamatRS
        .txtInstalasi.SetText "HUMAN RESOURCE DEPARTMENT"
        .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos : " & " " & strNKodepos & " "
        
        .UsNama.SetUnboundFieldSource ("{ado.NamaPegawai}")
        .usUnitKerja.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .ucInsentif.SetUnboundFieldSource ("{ado.Insentif}")
        .ucPph.SetUnboundFieldSource ("{ado.PPh}")
        .ucDiterima.SetUnboundFieldSource ("{ado.Diterima}")
        
        .unTotalIndex.SetUnboundFieldSource ("{ado.IndexPegawai}")
        .usBulanHitung.SetUnboundFieldSource ("{ado.Bulan}")
        .usTahunHitung.SetUnboundFieldSource ("{ado.Tahun}")
'        .SelectPrinter sDriver, sPrinter, vbNull
'         settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
        .DisplayGroupTree = False
    End With
    Screen.MousePointer = vbDefault
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strSQL = "delete from vJasaRemunerasiPegawai WHERE NamaKomputer = '" & strNamaHostLocal & "'  AND Bulan= '" & MonthName(Month(frmInsentifKaryawan.dtpPeriode.Value)) & "' and Tahun='" & Year(frmInsentifKaryawan.dtpPeriode.Value) & "'"
    dbConn.Execute strSQL
    Set frmCetakIndexPegawaiPerRuangan = Nothing
End Sub


