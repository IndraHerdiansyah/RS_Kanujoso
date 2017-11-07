VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSuratKeteranganPerjalananDinas 
   Caption         =   "Cetak Surat Keterangan Perjalanan Dinas"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakSuratKeteranganPerjalananDinas.frx":0000
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
Attribute VB_Name = "frmCetakSuratKeteranganPerjalananDinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crSuratKeteranganPerjalananDinas

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    On Error GoTo errLoad

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    
    Report.txtNamaKota.SetText strNKotaRS
    Report.txtNamaKota2.SetText strNKotaRS

    With Report

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        .Database.AddADOCommand dbConn, adocomd

        .txtNama.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
        .txtNip.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
        .txtJabatan.SetText IIf(IsNull(rs.Fields("Jabatan Pangkat").Value), "", rs.Fields("Jabatan Pangkat").Value)
        .txtKota.SetText IIf(IsNull(rs.Fields("KotaTujuan").Value), "", rs.Fields("KotaTujuan").Value)
        .txtkendaraan.SetText IIf(IsNull(rs.Fields("Kendaraan").Value), "", rs.Fields("Kendaraan").Value)
        .txttglberangkat.SetText IIf(IsNull(Format(rs.Fields("TglPergi").Value, "dd MMMM yyyy")), "", Format(rs.Fields("TglPergi").Value, "dd MMMM yyyy"))
        .txttglkembali.SetText IIf(IsNull(Format(rs.Fields("TglPulang").Value, "dd MMMM yyyy")), "", Format(rs.Fields("TglPulang").Value, "dd MMMM yyyy"))
        .txtPenyandangDana.SetText IIf(IsNull(rs.Fields("PenyandangDana").Value), "", rs.Fields("PenyandangDana").Value)
        .txtKeterangan.SetText IIf(IsNull(rs.Fields("Keterangan").Value), "", rs.Fields("Keterangan").Value)
        .txttujuan.SetText IIf(IsNull(rs.Fields("TujuanKunjungan").Value), "", rs.Fields("TujuanKunjungan").Value)
        .txtPemegang.SetText IIf(IsNull(rs.Fields("NamaLengkap").Value), "", rs.Fields("NamaLengkap").Value)
        .txtNipPemegang.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
        .txtAngkaHari.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", rs.Fields("Lamanya").Value)
        .txtHurufHari.SetText IIf(IsNull(rs.Fields("Lamanya").Value), "", NumToText(rs.Fields("Lamanya").Value))

        If strTandaTangan = "A01" Then
            .Text49.Suppress = True
            .Text23.Suppress = False
            strSQLsplakuk = "select NamaLengkap, NamaPangkat, NIP from V_FooterPegawai where KdJabatan='A01'"
            Call msubRecFO(rsSplakuk, strSQLsplakuk)

            .txtDirektur.SetText IIf(IsNull(rsSplakuk("NamaLengkap").Value), "", rsSplakuk("NamaLengkap").Value)
            .txtPangkatF.SetText IIf(IsNull(rsSplakuk("NamaPangkat").Value), "", rsSplakuk("NamaPangkat").Value)
            .txtNIPF.SetText IIf(IsNull(rsSplakuk("NIP").Value), "", rsSplakuk("NIP").Value)
        Else
            .Text23.Suppress = True
            .Text49.Suppress = False
            strSQLsplakuk = "select NamaLengkap, NamaPangkat, NIP from V_FooterPegawai where KdJabatan='" & strTandaTangan & "'"
            Call msubRecFO(rsSplakuk, strSQLsplakuk)

            .txtDirektur.SetText IIf(IsNull(rsSplakuk("NamaLengkap").Value), "", rsSplakuk("NamaLengkap").Value)
            .txtPangkatF.SetText IIf(IsNull(rsSplakuk("NamaPangkat").Value), "", rsSplakuk("NamaPangkat").Value)
            .txtNIPF.SetText IIf(IsNull(rsSplakuk("NIP").Value), "", rsSplakuk("NIP").Value)
        End If
    End With

    CRViewer1.ReportSource = Report
    CRViewer1.Zoom 1
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
    Exit Sub

errLoad:
    Call msubPesanError
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakSuratKeteranganPerjalananDinas = Nothing
End Sub
