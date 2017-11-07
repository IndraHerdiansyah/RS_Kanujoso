VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakRekapIndexRemunerasi 
   Caption         =   "Laporan Rekapitulasi Index Remunerasi"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5265
   Icon            =   "frmCetakRekapIndexsPegawai.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   5265
   WindowState     =   2  'Maximized
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
Attribute VB_Name = "frmCetakRekapIndexRemunerasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crRekapIndexRemunerasi

Private Sub Form_Load()
On Error GoTo errLoad
Dim i As Integer

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn

    strsql = "select  distinct B.KdRuanganKerja, B.NamaRuangan, B.NamaPegawai, B.NIP, B.Jabatan, B.IdPegawai, " & _
             "BasicInd=isnull(sum(A.BasicInd),0), CapaInd=isnull(sum(A.CapaInd),0), RiskInd=isnull(sum(A.RiskInd),0), EmergInd=isnull(sum(A.EmergInd),0),PositionInd=isnull(sum(A.PositionInd),0), PerforInd=isnull(sum(A.PerforInd),0), A.ThnTglHitung, A.BlnTglHitung " & _
             "from " & _
                "(select IdPegawai, BasicInd=isnull(sum(BasicInd),0), CapaInd=isnull(sum(CapaInd),0), RiskInd=isnull(sum(RiskInd),0), EmergInd=isnull(sum(EmergInd),0),PositionInd=isnull(sum(PositionInd),0), PerforInd=isnull(sum(PerforInd),0), ThnTglHitung, BlnTglHitung " & _
                "from V_RekapIndeksPegawaiX where ThnTglHitung BETWEEN '" & frmKriteriaLaporan.dtpBulanPenghitungan.Year & "' AND '" & frmKriteriaLaporan.dtpBulanPenghitunganAkhir.Year & "' and BlnTglHitung BETWEEN  " & _
                "'" & frmKriteriaLaporan.dtpBulanPenghitungan.Month & "' AND '" & frmKriteriaLaporan.dtpBulanPenghitunganAkhir.Month & "' " & _
                "group by IdPegawai, ThnTglHitung, BlnTglHitung) as A " & _
                "right outer join " & _
                "(SELECT DataPegawai.IdPegawai, ISNULL(DataCurrentPegawai.KdRuanganKerja, '---') AS KdRuanganKerja, ISNULL(RuangKerja.RuangKerja, '---') AS NamaRuangan, ISNULL(DataPegawai.NamaLengkap, '---') AS NamaPegawai, ISNULL(DataCurrentPegawai.NIP, '---') AS NIP, ISNULL(Jabatan.NamaJabatan, '---') AS Jabatan " & _
                    "FROM Jabatan RIGHT OUTER JOIN " & _
                         "RuangKerja RIGHT OUTER JOIN " & _
                         "DataCurrentPegawai LEFT OUTER JOIN " & _
                         "SubRuangKerja ON DataCurrentPegawai.KdRuanganKerja = SubRuangKerja.KdSubRuangKerja ON " & _
                         "RuangKerja.KdRuangKerja = SubRuangKerja.KdRuangKerja RIGHT OUTER JOIN " & _
                         "DataPegawai ON DataCurrentPegawai.IdPegawai = DataPegawai.IdPegawai ON Jabatan.KdJabatan = DataCurrentPegawai.KdJabatan " & _
                    "WHERE (DataCurrentPegawai.KdStatus = '01')) as B on A.idPegawai=B.idPegawai " & _
                    "where B.NamaRuangan like  '%" & frmKriteriaLaporan.DataCombo1.Text & "%' GROUP BY B.KdRuanganKerja, B.NamaRuangan, B.NamaPegawai, B.NIP, B.Jabatan, B.IdPegawai, A.ThnTglHitung, A.BlnTglHitung order by B.NamaPegawai"

    Call msubRecFO(rs, strsql)
    
    adocomd.CommandText = strsql
    adocomd.CommandType = adCmdText

    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .txtPeriode.SetText "PERIODE: " & UCase(Format(frmKriteriaLaporan.dtpBulanPenghitungan, "MMMM yyyy") & " S/D " & Format(frmKriteriaLaporan.dtpBulanPenghitunganAkhir, "MMMM yyyy"))
        .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
        .txtAlamat.SetText strNAlamatRS
        .txtAlamat2.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos : " & " " & strNKodepos & " "
        
        .Text1.SetText pubStrRuangan
        .usNamaPegawai.SetUnboundFieldSource ("{ado.NamaPegawai}")
        .usNIP.SetUnboundFieldSource ("{ado.NIP}")
        .usJabatan.SetUnboundFieldSource ("{ado.Jabatan}")
        .usBasIndex.SetUnboundFieldSource ("{ado.BasicInd}")
        .usCapIndex.SetUnboundFieldSource ("{ado.CapaInd}")
        .usRiskIndex.SetUnboundFieldSource ("{ado.RiskInd}")
        .usEmegIndex.SetUnboundFieldSource ("{ado.EmergInd}")
        .usPostnIndex.SetUnboundFieldSource ("{ado.PositionInd}")
        .usPerfIndex.SetUnboundFieldSource ("{ado.PerforInd}")
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
    Set frmCetakRekapIndexRemunerasi = Nothing
End Sub



