VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakStrukGajiPegawai 
   Caption         =   "Medifirrst2000 - Cetak Struk Gaji Pegawai"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5850
   Icon            =   "frmCetakStrukGajiPegawai.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   5850
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   5805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5805
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
Attribute VB_Name = "frmCetakStrukGajiPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rptGajiPegawai As crStrukGajiPegawai

Private Sub Form_Load()

    Dim totalpenerimaan As Currency
    Dim totalPotongan As Currency
    Dim totalBersih As Currency
    Me.WindowState = 2
    Set dbcmd = New ADODB.Command
    Set dbcmd.ActiveConnection = dbConn
    Set rptGajiPegawai = New crStrukGajiPegawai
    Screen.MousePointer = vbHourglass
    strSQL = "select idPegawai, [Nama Lengkap], Jabatan, Golongan,[Ruangan Kerja] as Ruangan from V_M_DataPegawaiNew where idPegawai='" & frmPembayaranGajiPegawai2.dcNamaPegawai.BoundText & "'"
    Call msubRecFO(rs, strSQL)
    With rptGajiPegawai
        .txtNamaRS.SetText strNNamaRS
        .txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtNama.SetText rs("Nama Lengkap")
        .txtPangkat.SetText rs("Jabatan")
        .txtidPegawai.SetText rs!idpegawai
        .txtRuangan.SetText rs("Jabatan") 'rs!Ruangan
        .txtRekening.SetText ""
        
    End With
    'strSQL = "Select idPegawai, KomponenGaji, Jumlah, komponenpotongangaji, jumlahpotongan from PembayaranGajiDanPotonganPegawai where idPegawai='" & frmPembayaranGajiPegawai2.dcNamaPegawai.BoundText & "' and tglPembayaran = '" & Format(frmPembayaranGajiPegawai2.dtpTglBayar.Value, "yyyy/MM/dd") & "'"
    strSQL = "select * from V_Penggajian  where  idpegawai='" & frmPembayaranGajiPegawai2.dcNamaPegawai.BoundText & "' and TglPembayaran ='" & Format(frmPembayaranGajiPegawai2.dtpTglBayar.Value, "yyyy/MM/dd") & "'"
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    rptGajiPegawai.Database.AddADOCommand dbConn, dbcmd
    rptGajiPegawai.usJenisGaji.SetUnboundFieldSource ("{ado.Jenis}")
    rptGajiPegawai.usNamaPenerimaan.SetUnboundFieldSource ("{ado.KomponenGaji}")
    rptGajiPegawai.usJmlPenerimaan.SetUnboundFieldSource ("{ado.Jumlah}")
'    rptGajiPegawai.usNamaPotongan.SetUnboundFieldSource ("{ado.KomponenPotonganGaji}")
'    rptGajiPegawai.usJmlPotongan.SetUnboundFieldSource ("{ado.JumlahPotongan}")

    strSQL = "SELECT SUM(Jumlah) As JumlahPenerimaan from PembayaranGajiPegawai where idPegawai='" & frmPembayaranGajiPegawai2.dcNamaPegawai.BoundText & "' and tglPembayaran = '" & Format(frmPembayaranGajiPegawai2.dtpTglBayar.Value, "yyyy/MM/dd") & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs("JumlahPenerimaan")) Then
        totalpenerimaan = 0
    Else
        totalpenerimaan = rs("JumlahPenerimaan").Value
    End If

    rptGajiPegawai.txtTotalPendapatan.SetText Format(totalpenerimaan, "##,###.00")
    strSQL = "SELECT SUM(Jumlah) As JumlahPotongan from PembayaranPotonganGajiPegawai where idPegawai='" & frmPembayaranGajiPegawai2.dcNamaPegawai.BoundText & "' and tglPembayaran = '" & Format(frmPembayaranGajiPegawai2.dtpTglBayar.Value, "yyyy/MM/dd") & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs("JumlahPotongan")) Then
        totalPotongan = 0
    Else
        totalPotongan = rs("JumlahPotongan")
    End If
    rptGajiPegawai.txtTotalPotongan.SetText Format(totalPotongan, "##,###.00")

    rptGajiPegawai.txtTotalBersih.SetText Format(totalpenerimaan - totalPotongan, "##,###.00")
    
    rptGajiPegawai.txtBulan.SetText Format(frmPembayaranGajiPegawai2.dtpTglBayar.Value, "mmmm yyyy")

    CRViewer1.ReportSource = rptGajiPegawai
    With CRViewer1
        .EnableGroupTree = False
        .ViewReport
        .Zoom 1
    End With

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    With CRViewer1
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakStrukGajiPegawai = Nothing
End Sub
