VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSuratTugas 
   Caption         =   "Cetak Surat Tugas"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakSuratTugas.frx":0000
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
Attribute VB_Name = "frmCetakSuratTugas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crSuratTugas

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    On Error GoTo Errload

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    
    Report.txtNamaKota.SetText strNKotaRS

    With Report
        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        .Database.AddADOCommand dbConn, adocomd

        .txtNama.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
        .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
        .txtJabatan.SetText IIf(IsNull(rs.Fields("NamaJabatan").Value), "", rs.Fields("NamaJabatan").Value)
        .txtPangkat.SetText IIf(IsNull(rs.Fields("Pangkat Golongan").Value), "", rs.Fields("Pangkat Golongan").Value)

        .UnboundDateTime1.SetUnboundFieldSource ("{ado.TglMulai}")
        .UnboundDateTime2.SetUnboundFieldSource ("{ado.TglAkhir}")

        .txtketerangan.SetText IIf(IsNull(rs.Fields("NamaTugas").Value), "", rs.Fields("NamaTugas").Value)
        .txtDasar.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
        .txtLokasi.SetText IIf(IsNull(rs.Fields("alamat").Value), "", rs.Fields("alamat").Value)

        strsqlx = "select NamaLengkap, NamaPangkat, NIP from V_FooterPegawai where KdJabatan='01022'"
        Call msubRecFO(rsx, strsqlx)
        
        If rsx.EOF = True Then
            .txtPenanggungJawab.Suppress = False
            .txtPenanggungJawab.SetText "Penanggung Jawab"
        Else
            .txtDirektur.SetText IIf(IsNull(rsx("NamaLengkap").Value), "", rsx("NamaLengkap").Value)
            .txtPangkatF.SetText IIf(IsNull(rsx("NamaPangkat").Value), "", rsx("NamaPangkat").Value)
            .txtNIPF.SetText IIf(IsNull(rsx("NIP").Value), "", rsx("NIP").Value)
        End If
    End With
    If vLaporan = "Print" Then
        Report.PrintOut False
        Unload Me
    Else
        CRViewer1.ReportSource = Report
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
    End If
    Exit Sub

Errload:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakSuratTugas = Nothing
End Sub
