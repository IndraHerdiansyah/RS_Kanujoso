VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_LaporanRealisasiKenaikanGaji 
   Caption         =   "Medifirst2000 - Cetak Laporan Realisasi Kenaikan Gaji Berkala"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   Icon            =   "frm_cetak_LaporanRealisasiKenaikanGaji.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7950
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
Attribute VB_Name = "frm_cetak_LaporanRealisasiKenaikanGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crLaporanRealisasiKenaikanGaji

Private Sub Form_Load()
    On Error GoTo errLoad
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Set frm_cetak_LaporanRealisasiKenaikanGaji = Nothing
    Dim adocomd As New ADODB.Command

    adocomd.ActiveConnection = dbConn

    adocomd.CommandText = strSQL

    adocomd.CommandType = adCmdText

    With Report
        .Database.AddADOCommand dbConn, adocomd

        .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaNip}")
        .usPangkat.SetUnboundFieldSource ("{ado.NamaGolongan}")
        .udTMT.SetUnboundFieldSource ("{ado.TMT}")
        .ucGaji.SetUnboundFieldSource ("{ado.GajiPokok}")
        .usJabatan.SetUnboundFieldSource ("{ado.NamaJabatan}")
        .udTglSKUsulan.SetUnboundFieldSource ("{ado.TglSKUsulan}")
        .usNoSKUsulan.SetUnboundFieldSource ("{ado.NoSKUsulan}")
        .udTMTUsulan.SetUnboundFieldSource ("{ado.TglBerlaku}")
        .ucGajiUsulan.SetUnboundFieldSource ("{ado.GajiPokokUsulan}")
        .udTglSKSetuju.SetUnboundFieldSource ("{ado.TglSKReal}")
        .usNoSKSetuju.SetUnboundFieldSource ("{ado.NoSKReal}")
        .udTMTSetuju.SetUnboundFieldSource ("{ado.TglBerlaku}")
        .ucGajiSetuju.SetUnboundFieldSource ("{ado.GajiPokokUsulan}")
        .udTglSKReal.SetUnboundFieldSource ("{ado.TglSKReal}")
        .usNoSKReal.SetUnboundFieldSource ("{ado.NoSKReal}")
        .ucGajiReal.SetUnboundFieldSource ("{ado.GajiPokokUsulan}")
        .udTMTReal.SetUnboundFieldSource ("{ado.TglBerlaku}")
        .usKeterangan.SetUnboundFieldSource ("{ado.KeteranganLainnya}")

    End With

    strSQLsplakuk = "select NamaLengkap, NamaPangkat, NIP from V_FooterPegawai where KdJabatan='A01'"
    Call msubRecFO(rsSplakuk, strSQLsplakuk)

    Report.txtDirektur.SetText IIf(IsNull(rsSplakuk("NamaLengkap").Value), "", rsSplakuk("NamaLengkap").Value)
    Report.txtNIPF.SetText IIf(IsNull(rsSplakuk("NIP").Value), "", rsSplakuk("NIP").Value)

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
    Set frm_cetak_LaporanRealisasiKenaikanGaji = Nothing
End Sub
