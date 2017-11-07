VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_LaporanDaftarPegawaiBerDasarkanGolongan 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm_cetak_LaporanDaftarPegawaiBgolongan.frx":0000
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
Attribute VB_Name = "frm_cetak_LaporanDaftarPegawaiBerDasarkanGolongan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crLaporanDafatarPegawaiBerGolongan

Private Sub Form_Load()
    On Error GoTo hell
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
  
    Dim adocomd As New ADODB.Command

    adocomd.ActiveConnection = dbConn

    adocomd.CommandText = strSQL

    adocomd.CommandType = adCmdText

    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtNmRs.SetText strNNamaRS
        .txtAlamatRs.SetText strNAlamatRS
        .txtNoTlp.SetText strNTeleponRS & " - Faks - " & strNTeleponRS
        .txtJudul.SetText "LAPORAN DAFTAR PEGAWAI RUMAH SAKIT MENURUT GOLONGAN"
        .txtPriode.SetText Format(frmLaporanDaftarPegawaiMenurutGolongan.DTPickerAwal.Value, "MMMM-YYYY")
        .usStatus.SetUnboundFieldSource ("{ado.TypePegawai}")
        .unS2.SetUnboundFieldSource ("{ado.gol1a}")
        .unS1.SetUnboundFieldSource ("{ado.gol1b}")
        .unD4.SetUnboundFieldSource ("{ado.gol1c}")
        .unD3.SetUnboundFieldSource ("{ado.gol1d}")
        .unD1.SetUnboundFieldSource ("{ado.gol2a}")
        .unSma.SetUnboundFieldSource ("{ado.gol2b}")
        .unSmk.SetUnboundFieldSource ("{ado.gol2c}")
        .unStm.SetUnboundFieldSource ("{ado.gol2d}")
        .unSpk.SetUnboundFieldSource ("{ado.gol3a}")
        .unSmp.SetUnboundFieldSource ("{ado.gol3b}")
        .unSdn.SetUnboundFieldSource ("{ado.gol3c}")
        .ungol3d.SetUnboundFieldSource ("{ado.gol3d}")
        .un4a.SetUnboundFieldSource ("{ado.gol4a}")
        .un4b.SetUnboundFieldSource ("{ado.gol4b}")
        .un4c.SetUnboundFieldSource ("{ado.gol4c}")
        .un4d.SetUnboundFieldSource ("{ado.gol4d}")
    End With
   

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
    Set frm_cetak_LaporanDaftarPegawaiBerDasarkanGolongan = Nothing
End Sub
