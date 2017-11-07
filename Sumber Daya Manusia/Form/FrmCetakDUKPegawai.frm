VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakDUKPegawai 
   Caption         =   "Medifirst2000 - Data Pegawai"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmCetakDUKPegawai.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
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
Attribute VB_Name = "FrmCetakDUKPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New CrDUKPegawai

Private Sub Form_Load()

    Me.Caption = "Medifirst2000 - Cetak Daftar Urut Kepangkatan ( DUK ) PNS"
    Set Report = New CrDUKPegawai

    'SELECT     IdPegawai, NIP, Nama, Golongan, TMTG, Jabatan, Pendidikan, Tanggal, Bulan, Tahun, NoUrutPangkat, NoUrutJabatan, NoUrutGolongan, TempatLahir, TglLahir, Usia,
                '[Alamat Lengkap], KdPendidikan, NoUrutAlamat, NoUrutPendidikan, ThnLulus, TGLMUTASI, SatuanKerja, TMTJ, NamaPendidikan, FakultasJurusan, TglLulus,
                'NamaGolongan , TMTP, Pangkat, TglMasuk, MasaKerjaThn, MasaKerjaBln
    'FROM         V_duknew2_1 WHERE     (Nama LIKE '%Abdul Jalal, AMK%') ORDER BY Nama
    dbcmd.CommandText = strSQL
    dbcmd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, dbcmd
    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtKabupaten.SetText strNAlamatRS
        .txtAlamatRS.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS

        .unNomor.SetUnboundFieldSource ("{Ado.no}")
        .UsNama.SetUnboundFieldSource ("{Ado.Nama}")
        .UsNIP.SetUnboundFieldSource ("{Ado.Nip}")
        .UsGolongan.SetUnboundFieldSource ("{Ado.Pangkat1}")
        .udTMTGol.SetUnboundFieldSource ("{Ado.Pangkat2}")
        .UsJabatan.SetUnboundFieldSource ("{Ado.Jabatan1}")
        .udTMTJabatan.SetUnboundFieldSource ("{Ado.Jabatan2}")
        .usMasakerjaThn.SetUnboundFieldSource ("{Ado.MasaKerja1}")
        .usMasakerjaBln.SetUnboundFieldSource ("{Ado.MasaKerja2}")
        .usNamaJabatan.SetUnboundFieldSource ("{Ado.LatihanJabatan1}")
        .usThnSKJabatan.SetUnboundFieldSource ("{Ado.LatihanJabatan2}")
        .usEvaluasiDiklat.SetUnboundFieldSource ("{Ado.LatihanJabatan3}")
        .usNamaPendidikan.SetUnboundFieldSource ("{Ado.Pendidikan1}")
        .usThnLulus.SetUnboundFieldSource ("{Ado.Pendidikan2}")
        .UsPendidikan.SetUnboundFieldSource ("{Ado.Pendidikan3}")
        .usUmur.SetUnboundFieldSource ("{Ado.Usia}")
        .usMutasiKerja.SetUnboundFieldSource ("{Ado.MutasiKerja}")
        
'        .UsNama.SetUnboundFieldSource ("{Ado.Nama}")
'        .udTglLahir.SetUnboundFieldSource ("{Ado.TglLahir}")
'        .UsPangkat.SetUnboundFieldSource ("{Ado.Pangkat}")
'        .UsGolongan.SetUnboundFieldSource ("{Ado.Golongan}")
'        .udTMTPangkat.SetUnboundFieldSource ("{Ado.TMTP}")
'        .udTMTGol.SetUnboundFieldSource ("{Ado.TMTG}")
'        .UsJabatan.SetUnboundFieldSource ("{Ado.Jabatan}")
'        .udTMTJabatan.SetUnboundFieldSource ("{Ado.TMTJ}")
'        .UsPendidikan.SetUnboundFieldSource ("{Ado.Pendidikan}")
'        .UsNIP.SetUnboundFieldSource ("{Ado.NIP}")
'        .unNoBulan.SetUnboundFieldSource ("{Ado.Bulan}")
'        .unNoGolongan.SetUnboundFieldSource ("{Ado.NoUrutGolongan}")
'        .unNoJabatan.SetUnboundFieldSource ("{Ado.NoUrutJabatan}")
'        .unNoPangkat.SetUnboundFieldSource ("{Ado.NoUrutPangkat}")
'        .unNoPendidikan.SetUnboundFieldSource ("{Ado.NoUrutPendidikan}")
'        .unNoTahun.SetUnboundFieldSource ("{Ado.Tahun}")
'        .unNoTanggal.SetUnboundFieldSource ("{Ado.Tanggal}")
'        .unUsia.SetUnboundFieldSource ("{Ado.Usia}")
'        .usMasakerjaThn.SetUnboundFieldSource ("{Ado.MasaKerjaThn}")
'        .usMasakerjaBln.SetUnboundFieldSource ("{Ado.MasaKerjaBln}")
'        .usNamaPendidikan.SetUnboundFieldSource ("{Ado.NamaPendidikan}")
'        .usThnLulus.SetUnboundFieldSource ("{Ado.ThnLulus}")
'        .usUmur.SetUnboundFieldSource ("{Ado.Usia}")
'        .usMutasiKerja.SetUnboundFieldSource ("{Ado.SatuanKerja}")
        '.usNamaJabatan.SetUnboundFieldSource ("{Ado.NamaDiklat}")
        '.usThnSKJabatan.SetUnboundFieldSource ("{Ado.TglDiklat}")

    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "Print" Then
        Report.PrintOut False
        Unload Me
    Else
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .EnableGroupTree = False
            .Zoom 100

        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmCetakDUKPegawai = Nothing
End Sub
