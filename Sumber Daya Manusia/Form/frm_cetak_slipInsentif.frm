VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_cetak_SlipInsentif 
   Caption         =   "Medifirst2000 - Cetak Laporan Tahunan Usulan Pengangkatan TPHL"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   Icon            =   "frm_cetak_slipInsentif.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
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
Attribute VB_Name = "frm_cetak_SlipInsentif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crLaporanSlipInsentif
Public tgl As String

Private Sub Form_Load()
    On Error GoTo errLoad
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Set frm_cetak_SlipInsentif = Nothing
    Dim adocomd As New ADODB.Command

    adocomd.ActiveConnection = dbConn

    adocomd.CommandText = strSQL

    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd
    With Report
        .tgl.SetText tgl
    'nik, Nama, Jenis, Jabatan, Golongan, RJ, RI, m_RJ, m_RI, JasaTindakan, Kompentensi, Bruto, PPH21, InsentifBersih, Tjabatan
        .usNamaDokter.SetUnboundFieldSource ("{ado.Nama}")
        .usPosisi.SetUnboundFieldSource ("{ado.Jabatan}")
        .usStatus.SetUnboundFieldSource ("{ado.Jenis}")
        .usGol.SetUnboundFieldSource ("{ado.Golongan}")
        .usKompetensi.SetUnboundFieldSource ("{ado.Kompentensi}")
        .usJab.SetUnboundFieldSource ("{ado.Tjabatan}")
        .ucRawatJalan.SetUnboundFieldSource ("{ado.m_RJ}")
        .ucRanap.SetUnboundFieldSource ("{ado.m_RI}")
        .ucTindakan.SetUnboundFieldSource ("{ado.JasaTindakan}")
        .unRJ.SetUnboundFieldSource ("{ado.RJ}")
        .unRI.SetUnboundFieldSource ("{ado.RI}")
        .ucTotal.SetUnboundFieldSource ("{ado.Bruto}")
        .ucPajak.SetUnboundFieldSource ("{ado.PPH21}")
        .ucBersih.SetUnboundFieldSource ("{ado.InsentifBersih}")


    End With

'    strSQLsplakuk = "select NamaLengkap, NamaPangkat, NIP from V_FooterPegawai where KdJabatan='A01'"
'    Call msubRecFO(rsSplakuk, strSQLsplakuk)
'
'    Report.txtDirektur.SetText IIf(IsNull(rsSplakuk("NamaLengkap").Value), "", rsSplakuk("NamaLengkap").Value)
'    Report.txtNIPF.SetText IIf(IsNull(rsSplakuk("NIP").Value), "", rsSplakuk("NIP").Value)

'    Screen.MousePointer = vbHourglass

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
    Set frm_cetak_SlipInsentif = Nothing
End Sub
