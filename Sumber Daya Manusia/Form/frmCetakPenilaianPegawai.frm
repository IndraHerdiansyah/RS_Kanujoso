VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPenilaianPegawai 
   Caption         =   "Form Cetak Penilaian Pegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakPenilaianPegawai.frx":0000
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
Attribute VB_Name = "frmCetakPenilaianPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crPenilaianPegawai

Private Sub Form_Load()
    On Error GoTo hell
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    With Report
        .txtBulanAwal.SetText Format(rs.Fields("TglAwal").Value, "MMMM")
        .txtBulanAkhir.SetText Format(rs.Fields("TglAkhir").Value, "MMMM")
        .txtTahun.SetText Format(rs.Fields("TglAkhir").Value, "yy")

        .txtNamaPegawai.SetText rs.Fields("NamaLengkap").Value
        .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
        .txtTTDPegawai1.SetText rs.Fields("NamaLengkap").Value
        .txtTTDNIPPegawai1.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
        .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat Golongan").Value), "", rs.Fields("Pangkat Golongan").Value)
        .txtJabatan.SetText IIf(IsNull(rs.Fields("NamaJabatan").Value), "", rs.Fields("NamaJabatan").Value)

        .txtNamaPegawai2.SetText IIf(IsNull(rs.Fields("Penilai").Value), "", rs.Fields("Penilai").Value)
        .txtNIP2.SetText IIf(IsNull(rs.Fields("NIPPenilai").Value), "", rs.Fields("NIPPenilai").Value)
        .txtTTDPegawai2.SetText IIf(IsNull(rs.Fields("Penilai").Value), "", rs.Fields("Penilai").Value)
        .txtTTDNIPPegawai2.SetText IIf(IsNull(rs.Fields("NIPPenilai").Value), "", rs.Fields("NIPPenilai").Value)
        .txtPangkat2.SetText IIf(IsNull(rs.Fields("Pangkat Golongan Penilai").Value), "", rs.Fields("Pangkat Golongan Penilai").Value)
        .txtJabatan2.SetText IIf(IsNull(rs.Fields("JabatanPenilai").Value), "", rs.Fields("JabatanPenilai").Value)

        .txtNamaPegawai3.SetText IIf(IsNull(rs.Fields("AtasanPenilai").Value), "", rs.Fields("AtasanPenilai").Value)
        .txtNIP3.SetText IIf(IsNull(rs.Fields("NIPAtasan").Value), "", rs.Fields("NIPAtasan").Value)
        .txtTTDPegawai3.SetText IIf(IsNull(rs.Fields("AtasanPenilai").Value), "", rs.Fields("AtasanPenilai").Value)
        .txtTTDNIPPegawai3.SetText IIf(IsNull(rs.Fields("NIPAtasan").Value), "", rs.Fields("NIPAtasan").Value)
        .txtPangkat3.SetText IIf(IsNull(rs.Fields("Pangkat Golongan Atasan").Value), "", rs.Fields("Pangkat Golongan Atasan").Value)
        .txtJabatan3.SetText IIf(IsNull(rs.Fields("JabatanAtasan").Value), "", rs.Fields("JabatanAtasan").Value)
    End With

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 1
            .DisplayGroupTree = False
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
hell:
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakPenilaianPegawai = Nothing
End Sub
