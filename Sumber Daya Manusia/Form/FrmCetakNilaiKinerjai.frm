VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakNilaiKinerja 
   Caption         =   "Medifirst2000 - Daftar Penghasilan Pegawai Non PNS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmCetakNilaiKinerjai.frx":0000
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
      DisplayGroupTree=   0   'False
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
Attribute VB_Name = "FrmCetakNilaiKinerja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New CrNilaiKinerja

Private Sub Form_Load()
    On Error GoTo hell
    Dim adocomd As New ADODB.Command

    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn

    adocomd.CommandText = strSQL
    adocomd.CommandType = adCmdText

    Report.Database.AddADOCommand dbConn, adocomd
    With Report
        .txtNamaRS.SetText strNNamaRS & " " & strkelasRS & " " & strketkelasRS
        .txtKabupaten.SetText "KABUPATEN " & strNKotaRS
        .txtAlamatRS.SetText strNAlamatRS
'        .txtTelpRS.SetText "TELP : " & strNTeleponRS & " " & "Kode Pos : " & " " & strNKodepos & " "

        .txtPeriode.SetText "Bulan : " & Format(frmKinerja_Transaksi.DTPicker2.Value, "MMMM yyyy") & ""

        '.UsNama.SetUnboundFieldSource ("{Ado.NamaLengkap}")
        '.ucGapok.SetUnboundFieldSource ("{ado.GajiPokok}")
        '.unJmlAnak.SetUnboundFieldSource ("{ado.JmlAnak}")
        .UnboundNumber1.SetUnboundFieldSource ("{Ado.No}")
        .UnboundString2.SetUnboundFieldSource ("{Ado.NamaKinerja}")
        .UnboundNumber2.SetUnboundFieldSource ("{Ado.Nilai}")
        .UnboundNumber3.SetUnboundFieldSource ("{Ado.No2}")
        .UnboundString5.SetUnboundFieldSource ("{Ado.NamaKinerja2}")
        .UnboundNumber4.SetUnboundFieldSource ("{Ado.Nilai2}")
        
        Dim BHU, BPU As Integer
        Dim jmlBHU, jmlBPU As Integer
        strsqlx = "SELECT    SUM( Nilai),sum(Nilai2), kode FROM         tempCetakKinerjaBulan where kode='" & frmKinerja_Transaksi.txtidpegawai.Text & "' group by kode"
        Call msubRecFO(rsAplikasi, strsqlx)
        If rsAplikasi.RecordCount <> 0 Then
            .txtTotalBHU.SetText rsAplikasi(0)
            .txtTotalBPU.SetText rsAplikasi(1)
            BHU = rsAplikasi(0)
            BPU = rsAplikasi(1)
        End If
        strsqlx = "SELECT  MasterKinerja.KdKategoryKinerja,  COUNT( NilaiKinerjaPegawai.KdKinerja) " & _
                  "FROM         NilaiKinerjaPegawai INNER JOIN " & _
                  "MasterKinerja ON NilaiKinerjaPegawai.KdKinerja = MasterKinerja.KdKinerja " & _
                  "WHERE     (NilaiKinerjaPegawai.IdPegawai = '" & frmKinerja_Transaksi.txtidpegawai.Text & "') and NilaiKinerjaPegawai.bulan ='" & Format(frmKinerja_Transaksi.DTPicker2.Value, "MM") & "' and NilaiKinerjaPegawai.tahun='" & Format(frmKinerja_Transaksi.DTPicker2.Value, "yyyy") & "'" & _
                  "group by MasterKinerja.KdKategoryKinerja"
        Call msubRecFO(rsAplikasi, strsqlx)
        If rsAplikasi.RecordCount <> 0 Then
            For i = 0 To rsAplikasi.RecordCount - 1
                If rsAplikasi(0) = "01" Then jmlBHU = rsAplikasi(1)
                If rsAplikasi(0) = "02" Then jmlBPU = rsAplikasi(1)
                rsAplikasi.MoveNext
            Next
        End If
        Dim NilaiTotalPresiden As Double
        .txtBHU.SetText FormatNumber(((BHU / jmlBHU) * 70) / 100, 2)
        .txtBPU.SetText FormatNumber(((BPU / jmlBPU) * 30) / 100, 2)
        NilaiTotalPresiden = FormatNumber((((BHU / jmlBHU) * 70) / 100) + (((BPU / jmlBPU) * 30) / 100), 2)
        .txtGrandTotal.SetText NilaiTotalPresiden
        
        Dim NilaiPresiden As String
        If NilaiTotalPresiden > 90 Then NilaiPresiden = "Sangat Baik"
        If NilaiTotalPresiden > 76 And NilaiTotalPresiden <= 90 Then NilaiPresiden = "Baik"
        If NilaiTotalPresiden > 61 And NilaiTotalPresiden <= 75 Then NilaiPresiden = "Cukup"
        If NilaiTotalPresiden > 51 And NilaiTotalPresiden <= 60 Then NilaiPresiden = "Kurang"
        If NilaiTotalPresiden <= 50 Then NilaiPresiden = "Buruk"
        .txtHasilPenilaian.SetText NilaiPresiden
        
        strSQL = "delete from RekapKinerjaBulan WHERE     (IdPegawai = '" & frmKinerja_Transaksi.txtidpegawai.Text & "') and bulan ='" & Format(frmKinerja_Transaksi.DTPicker2.Value, "MM") & "' and tahun='" & Format(frmKinerja_Transaksi.DTPicker2.Value, "yyyy") & "'"
        Call msubRecFO(rs, strSQL)
        strSQL = "insert into RekapKinerjaBulan values('" & frmKinerja_Transaksi.txtidpegawai.Text & "' ,'" & Format(frmKinerja_Transaksi.DTPicker2.Value, "MM") & "' ,'" & Format(frmKinerja_Transaksi.DTPicker2.Value, "yyyy") & "','" & NilaiTotalPresiden & "','" & NilaiPresiden & "')"
       Call msubRecFO(rs, strSQL)
        
        strSQL = "SELECT     DataPegawai.IdPegawai, DataPegawai.NamaLengkap, DataCurrentPegawai.NIP, Jabatan.NamaJabatan, DataCurrentPegawai.KdPegawaiAtasan, GolonganPegawai.NamaGolongan " & _
                 "FROM         Pangkat INNER JOIN " & _
                 "DataCurrentPegawai ON Pangkat.KdPangkat = DataCurrentPegawai.KdPangkat INNER JOIN " & _
                 "GolonganPegawai ON Pangkat.KdGolongan = GolonganPegawai.KdGolongan LEFT OUTER JOIN " & _
                 "Jabatan ON DataCurrentPegawai.KdJabatan = Jabatan.KdJabatan RIGHT OUTER JOIN " & _
                 "DataPegawai ON DataCurrentPegawai.IdPegawai = DataPegawai.IdPegawai " & _
                 " where DataPegawai.IdPegawai='" & frmKinerja_Transaksi.txtidpegawai.Text & "'"
        Call msubRecFO(rsc, strSQL)
        .a.SetText IIf(IsNull(rsc(1)), "", rsc(1))
        .a1.SetText IIf(IsNull(rsc(2)), "", rsc(2)) 'rsc(2)
        .a2.SetText IIf(IsNull(rsc(3)), "", rsc(3)) 'rsc(3)
        .txtttdPegawai.SetText IIf(IsNull(rsc(1)), "", rsc(1))
        
        If IsNull(rsc(4)) = False Then
            strSQL = "SELECT     DataPegawai.IdPegawai, DataPegawai.NamaLengkap, DataCurrentPegawai.NIP, Jabatan.NamaJabatan, DataCurrentPegawai.KdPegawaiAtasan, GolonganPegawai.NamaGolongan " & _
                 "FROM         Pangkat INNER JOIN " & _
                 "DataCurrentPegawai ON Pangkat.KdPangkat = DataCurrentPegawai.KdPangkat INNER JOIN " & _
                 "GolonganPegawai ON Pangkat.KdGolongan = GolonganPegawai.KdGolongan LEFT OUTER JOIN " & _
                 "Jabatan ON DataCurrentPegawai.KdJabatan = Jabatan.KdJabatan RIGHT OUTER JOIN " & _
                 "DataPegawai ON DataCurrentPegawai.IdPegawai = DataPegawai.IdPegawai " & _
                 " where DataPegawai.IdPegawai='" & rsc(4) & "'"
            Call msubRecFO(rsb, strSQL)
            .b1.SetText IIf(IsNull(rsb(1)), "", rsb(1)) 'rsb(1)
            .b2.SetText IIf(IsNull(rsb(2)), "", rsb(2)) 'rsb(2)
            .b3.SetText IIf(IsNull(rsb(5)), "", rsb(5)) 'rsb(2)
            .b4.SetText IIf(IsNull(rsb(3)), "", rsb(3)) 'rsb(3)
            .txtttdAtasan.SetText IIf(IsNull(rsb(1)), "", rsb(1)) 'rsb(1)
        End If
    End With
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
    Set FrmCetakNilaiKinerja = Nothing
End Sub
