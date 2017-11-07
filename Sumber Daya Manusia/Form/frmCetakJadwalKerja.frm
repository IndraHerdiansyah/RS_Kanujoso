VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakJadwalKerja 
   Caption         =   "frmCetakDaftarNoAbsensi"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakJadwalKerja.frx":0000
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
Attribute VB_Name = "frmCetakJadwalKerja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crJadwalKerja

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn

    If strCetak = "CetakJadwal" Then
'        strSQL = "SELECT Nama, Ruangan, Shift, Tanggal, Singkatan From v_JadwalKerjaNew " & _
        " Where Ruangan LIKE '%" & strGroup & "%' " & _
        " AND Tanggal BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd") & "' "
        '// yayang.agus 2014-08-11
        strSQL = "SELECT Nama, Ruangan, Shift, Tanggal From v_JadwalKerjaNew " & _
        " Where Ruangan LIKE '%" & strGroup & "%' " & _
        " AND Tanggal BETWEEN '" & Format(mdTglAwal, "yyyy/MM/dd") & "' AND '" & Format(mdTglAkhir, "yyyy/MM/dd") & "' "
        '//

        Report.txtRuangan.SetText strGroup

        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        Report.Database.AddADOCommand dbConn, adocomd
        With Report
            .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
            .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
            .txtAlamatRS.SetText strNAlamatRS
            .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos : " & " " & strNKodepos & " "
            .txtPeriode.SetText "Periode : " & Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy") & ""

            .usNamaPegawai.SetUnboundFieldSource ("{Ado.Nama}")
            '.usShift.SetUnboundFieldSource ("{Ado.Singkatan}")
            .usShift.SetUnboundFieldSource ("{Ado.Shift}") '//yayang.agus 2014-08-11
            .udTanggal.SetUnboundFieldSource ("{Ado.Tanggal}")

        End With
    End If
    With CRViewer1
        .ReportSource = Report
        .EnableGroupTree = False
        .EnableExportButton = True
        .ViewReport
        .Zoom 100
    End With

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakJadwalKerja = Nothing
End Sub

