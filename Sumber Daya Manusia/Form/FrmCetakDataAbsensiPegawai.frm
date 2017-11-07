VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakDataAbsensiPegawai 
   Caption         =   "Medifirst2000 - Data Pegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmCetakDataAbsensiPegawai.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
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
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmCetakDataAbsensiPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrPegawai

Private Sub Form_Load()
Dim adocomd As New ADODB.Command

    Call openConnection
    Set frmDataPegawai = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "SELECT     IdPegawai, JenisPegawai, NamaLengkap, JenisKelamin, TempatLahir, TglLahir, NamaPangkat, NamaGolongan, NamaJabatan, Pendidikan, NIP, StatusAktif " & _
    " From V_Data_Pegawai1 "
    
      adocomd.CommandType = adCmdUnknown
    Report.Database.AddADOCommand dbConn, adocomd
    With Report
    .usIDPegawai.SetUnboundFieldSource ("{Ado.IdPegawai}")
    .usJenisPegawai.SetUnboundFieldSource ("{Ado.JenisPegawai}")
    .UsNama.SetUnboundFieldSource ("{Ado.NamaLengkap}")
    .UsJK.SetUnboundFieldSource ("{Ado.JenisKelamin}")
    .usTempatLahir.SetUnboundFieldSource ("{Ado.TempatLahir}")
    .UsTanggallahir.SetUnboundFieldSource ("{Ado.Tgllahir}")
    .usPangkat.SetUnboundFieldSource ("{Ado.namaPangkat}")
    .usGolongan.SetUnboundFieldSource ("{Ado.namaGolongan}")
    .UsPendidikasn.SetUnboundFieldSource ("{Ado.Pendidikan}")
    .usStatus.SetUnboundFieldSource ("{Ado.StatusAktif}")
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .txtWebsiteRS.SetText strWebsite & ", " & strEmail
        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .EnableGroupTree = True
        .Zoom 1
        
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set FrmCetakDataPegawai = Nothing
End Sub


