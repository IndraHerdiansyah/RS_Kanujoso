VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSKUM 
   Caption         =   "Form Cetak Surat Keterangan  Untuk Mendapatkan Pembayaran Tunjangan Keluarga"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCetakSKUM.frx":0000
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
Attribute VB_Name = "frmCetakSKUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crSKUM

Private Sub Form_Load()
    On Error GoTo hell
    Dim adocomd As New ADODB.Command
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    strSQL = "Select * from V_SKUM where IdPegawai = '" & mstrIdPegawai & "' "
    Call msubRecFO(rs, strSQL)
    With Report
        adocomd.CommandText = strSQL
        adocomd.CommandType = adCmdText
        .Database.AddADOCommand dbConn, adocomd

        .txtnamapegawai.SetText rs.Fields("NamaLengkap").Value
        .txtNIP.SetText IIf(IsNull(rs.Fields("NIP").Value), "", rs.Fields("NIP").Value)
        .txtPangkatGol.SetText rs.Fields("NamaPangkat").Value + "/" + rs.Fields("NamaGolongan").Value
        .txtTMTGol.SetText Format(rs.Fields("TglSKGol").Value, "dd MMMM yyyy")
        .txtTTL.SetText rs.Fields("TempatLahir").Value + "/" + Format(rs.Fields("TglLahir").Value, "dd MMMM yyyy")
        .txtJKStatus.SetText rs.Fields("JenisKelamin").Value + "/" + rs.Fields("StatusPerkawinan").Value
        .txtAgama.SetText rs.Fields("Agama").Value + "/" + "Indonesia"
        .txtAlamat.SetText rs.Fields("AlamatLengkap").Value
        .txtRTRW.SetText rs.Fields("RTRW").Value
        .txtDesa.SetText rs.Fields("Kelurahan").Value
        .txtKecamatan.SetText rs.Fields("Kecamatan").Value
        .txtKota.SetText rs.Fields("KotaKabupaten").Value
        .txtTMTCapeg.SetText Format(rs.Fields("TglMasuk").Value, "dd MMMM yyyy")
        .txtSKGaji.SetText rs.Fields("SKGaji").Value
        .txtSKTerakhir.SetText rs.Fields("NoSK").Value
        .txtJabatanF.SetText rs.Fields("JenisJabatan").Value

        strSQLsplakuk = "Select COUNT(NoUrut) as Jmltanggungan from KeluargaPegawai where IdPegawai = '" & mstrIdPegawai & "' "
        Call msubRecFO(rsSplakuk, strSQLsplakuk)
        .txtJmlTanggungan.SetText rsSplakuk.Fields("JmlTanggungan")

        strsqlx = "SELECT dbo.DataPegawai.NamaLengkap FROM dbo.DataPegawai INNER JOIN " & _
        "dbo.DataCurrentPegawai ON dbo.DataPegawai.IdPegawai = dbo.DataCurrentPegawai.IdPegawai " & _
        "WHERE (dbo.DataCurrentPegawai.KdJabatan = 'A04') "
        Call msubRecFO(rsx, strsqlx)
        .txtBendaharawan.SetText rsx.Fields(0).Value

        .txtNamaFoot.SetText rs.Fields("NamaLengkap").Value
        .txtNIPFoot.SetText rs.Fields("NIP").Value
    End With

    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
        .DisplayGroupTree = False
    End With
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
    Set frmCetakSKUM = Nothing
End Sub
