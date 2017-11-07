VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRekapKinerja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medifirst2000 - Laporan Abensi Pegawai"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRekapKinerja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   10575
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdcetakdetail 
         Caption         =   "Cetak &Detail Absensi"
         Height          =   495
         Left            =   5760
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Cetak &Rekap"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periode :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   10575
      Begin VB.OptionButton optTahun 
         Caption         =   "Per Tahun"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton optJam 
         Caption         =   "Per Jam"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   720
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optHari 
         Caption         =   "Per Hari"
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optBulan 
         Caption         =   "Per Bulan"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   330
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   105971715
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   6000
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy HH:mm"
         Format          =   105971715
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   5520
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Group Cetak Berdasarkan :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   10575
      Begin VB.OptionButton optInstalasi 
         Caption         =   "Instalasi"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkGroup 
         Caption         =   "Group yang dipilih"
         Height          =   255
         Left            =   6840
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton optJenisPegawai 
         Caption         =   "Jenis Pegawai"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optPegawai 
         Caption         =   "Nama Pegawai"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optRuangan 
         Caption         =   "Ruangan"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optJabatan 
         Caption         =   "Jabatan"
         Height          =   375
         Left            =   5520
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcGroup 
         Height          =   390
         Left            =   6840
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   688
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   1440
      Picture         =   "frmRekapKinerja.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRekapKinerja.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRekapKinerja.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmRekapKinerja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    On Error GoTo errLoad
    
    'strSQL = " SELECT * " & _
    " FROM V_LaporanPengambilanARV " & _
    " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' "
    strSQL = "SELECT     RekapKinerjaBulan.idPegawai, DataPegawai.NamaLengkap, RekapKinerjaBulan.Bulan, RekapKinerjaBulan.Tahun, RekapKinerjaBulan.NilaiTotal, " & _
                      "RekapKinerjaBulan.NilaiPrestasi " & _
                    "FROM         RekapKinerjaBulan INNER JOIN " & _
                     "DataPegawai ON RekapKinerjaBulan.idPegawai = DataPegawai.IdPegawai " & _
                     "WHERE     RekapKinerjaBulan.Bulan ='" & Format(dtpAwal.Value, "MM") & "' "
                    '"WHERE     (RekapKinerjaBulan.Bulan between '" & Format(dtpAwal.Value, "MM") & "' AND '" & Format(dtpAkhir.Value, "MM") & "') "

   
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        Exit Sub
    End If
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value

    
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakRekapKinerja.Show
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.Value = Format(Now, "dd MMMM yyyy")
    dtpAkhir.Value = Format(Now, "dd MMMM yyyy")

End Sub


