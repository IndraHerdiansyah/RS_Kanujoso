VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLaporanDetailAbsensi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medifirst2000 - Laporan Abensi Pegawai"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLaporanDetailAbsensi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   3360
      Width           =   10575
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   8280
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdcetakdetail 
         Caption         =   "Cetak &Detail Absensi"
         Height          =   495
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdCetakAbsensi 
         Caption         =   "Cetak &Absensi"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2055
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
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   10575
      Begin VB.OptionButton optTahun 
         Caption         =   "Per Tahun"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   480
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
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   330
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
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
         Format          =   83558403
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   6000
         TabIndex        =   8
         Top             =   360
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
         Format          =   83558403
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
      Left            =   8760
      Picture         =   "frmLaporanDetailAbsensi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLaporanDetailAbsensi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanDetailAbsensi.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmLaporanDetailAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkGroup_Click()
    If chkGroup.Value = vbChecked Then
        dcGroup.Enabled = True
    Else
        dcGroup.Enabled = False
    End If
End Sub

Private Sub chkGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkGroup.Value = vbChecked Then
            dcGroup.SetFocus
        Else
            cmdCetakAbsensi.SetFocus
        End If
    End If
End Sub

Private Sub cmdCetakAbsensi_Click()
    On Error GoTo hell
    Dim pesan As VbMsgBoxResult
    strGroup = ""
    If optHari.Value = True Then
        strCetak = "Hari"
    ElseIf optBulan.Value = True Then
        strCetak = "Bulan"
    ElseIf optJam.Value = True Then
        strCetak = "Jam"
    Else
        strCetak = "Tahun"
    End If

    mdTglAwal = dtpAwal.Value 'TglAwal
    mdTglAkhir = dtpAkhir.Value 'TglAkhir

    If optInstalasi.Value = True Then
        strGroup = "NamaInstalasi"
    ElseIf optRuangan.Value = True Then
        strGroup = "NamaRuangan"
    ElseIf optPegawai.Value = True Then
        strGroup = "NamaLengkap"
    ElseIf optJenisPegawai.Value = True Then
        strGroup = "JenisPegawai"
    ElseIf optJabatan.Value = True Then
        strGroup = "Jabatan"
    Else
        strGroup = vbNullString
    End If

    If chkGroup.Value = vbChecked Then
        strIsiGroup = dcGroup.Text
    Else
        strIsiGroup = vbNullString
    End If

    strCetak2 = "CetakAbsensi"
    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"
    frmCetakAbsensiPegawai.Show
    Exit Sub
hell:
End Sub

Private Sub cmdcetakdetail_Click()
    On Error GoTo hell
    Dim pesan As VbMsgBoxResult
    strGroup = ""

    If optHari.Value = True Then
        strCetak = "Hari"
    ElseIf optBulan.Value = True Then
        strCetak = "Bulan"
    ElseIf optJam.Value = True Then
        strCetak = "Jam"
    ElseIf optTahun.Value = True Then
        strCetak = "Tahun"
    End If

    mdTglAwal = dtpAwal.Value 'TglAwal
    mdTglAkhir = dtpAkhir.Value 'TglAkhir

    If optInstalasi.Value = True Then
        strGroup = "NamaInstalasi"
    ElseIf optRuangan.Value = True Then
        strGroup = "NamaRuangan"
    ElseIf optPegawai.Value = True Then
        strGroup = "NamaLengkap"
    ElseIf optJenisPegawai.Value = True Then
        strGroup = "JenisPegawai"
    ElseIf optJabatan.Value = True Then
        strGroup = "Jabatan"
    Else
        strGroup = vbNullString
    End If

    If chkGroup.Value = vbChecked Then
        strIsiGroup = dcGroup.Text
    Else
        strIsiGroup = vbNullString
    End If

    strCetak2 = "CetakDetailAbsensi"

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    frmCetakAbsensiPegawai.Show
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcGroup_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcGroup.BoundText
    If optInstalasi.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdInstalasi, NamaInstalasi FROM Instalasi ORDER BY NamaInstalasi")
    ElseIf optRuangan.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan ORDER BY NamaRuangan")
    ElseIf optPegawai.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT IdPegawai, NamaLengkap FROM DataPegawai ORDER BY NamaLengkap")
    ElseIf optJenisPegawai.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdJenisPegawai, JenisPegawai FROM JenisPegawai ORDER BY JenisPegawai")
    ElseIf optJabatan.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdJabatan, NamaJabatan FROM Jabatan ORDER BY NamaJabatan")
    Else
        Exit Sub
    End If
    dcGroup.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcGroup_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcGroup.Text)) = 0 Then dcGroup.SetFocus: Exit Sub

        If optInstalasi.Value = True Then
            strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE (NamaInstalasi LIKE '%" & dcGroup.Text & "%')"
        ElseIf optRuangan.Value = True Then
            strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE (NamaRuangan LIKE '%" & dcGroup.Text & "%')"
        ElseIf optPegawai.Value = True Then
            strSQL = "SELECT IdPegawai, NamaLengkap FROM DataPegawai WHERE (NamaLengkap LIKE '%" & dcGroup.Text & "%')"
        ElseIf optJenisPegawai.Value = True Then
            strSQL = "SELECT KdJenisPegawai, JenisPegawai FROM JenisPegawai WHERE (JenisPegawai LIKE '%" & dcGroup.Text & "%')"
        ElseIf optJabatan.Value = True Then
            strSQL = "SELECT KdJabatan, NamaJabatan FROM Jabatan WHERE (NamaJabatan LIKE '%" & dcGroup.Text & "%')"
        Else
            optRuangan.SetFocus
        End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcGroup.BoundText = rs(0).Value
        dcGroup.Text = rs(1).Value
        cmdcetakdetail.SetFocus
    End If
    Exit Sub
hell:
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetakAbsensi.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    optRuangan.Value = True
    optBulan.Value = True
End Sub

Private Sub optInstalasi_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optInstalasi.Caption
End Sub

Private Sub optInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optTahun_Click()
    dtpAwal.CustomFormat = "yyyy"
    dtpAkhir.CustomFormat = "yyyy"
End Sub

Private Sub optJam_Click()
    dtpAwal.CustomFormat = "dd MMMM yyyy HH:mm"
    dtpAkhir.CustomFormat = "dd MMMM yyyy HH:mm"
End Sub

Private Sub optJenisPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optRuangan_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optRuangan.Caption
End Sub

Private Sub optRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optBulan_Click()
    dtpAwal.CustomFormat = "MMMM yyyy"
    dtpAkhir.CustomFormat = "MMMM yyyy"
End Sub

Private Sub optBulan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optPegawai_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optPegawai.Caption
End Sub

Private Sub optHari_Click()
    dtpAwal.CustomFormat = "dd MMMM yyyy"
    dtpAkhir.CustomFormat = "dd MMMM yyyy"
End Sub

Private Sub optHari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optJenisPegawai_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optJenisPegawai.Caption
End Sub

Private Sub optJabatan_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optJabatan.Caption
End Sub

Private Sub optJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub
