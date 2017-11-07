VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLaporanSaldoBarang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medifirst 2000 - Laporan Absensi Pegawai"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLaporanSaldoBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Periode"
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
      TabIndex        =   13
      Top             =   960
      Width           =   9255
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4335
         Begin VB.OptionButton optHari 
            Caption         =   "Per Hari"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optBulan 
            Caption         =   "Per Bulan"
            Height          =   375
            Left            =   1200
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   330
         Left            =   4560
         TabIndex        =   17
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58785795
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   7080
         TabIndex        =   18
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58785795
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   6720
         TabIndex        =   19
         Top             =   525
         Width           =   255
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   3480
      Width           =   9255
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   2760
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkSemuaRuangan 
         Caption         =   "Semua Ruangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdSpreadSheet 
         Caption         =   "&SpreadSheet"
         Height          =   495
         Left            =   4920
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   495
         Left            =   7080
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Group Cetak"
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
      TabIndex        =   9
      Top             =   2160
      Width           =   9255
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   5175
         Begin VB.OptionButton optJabatan 
            Caption         =   "Jabatan"
            Height          =   375
            Left            =   3960
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optRuangan 
            Caption         =   "Ruangan"
            Height          =   375
            Left            =   240
            TabIndex        =   0
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optInstalasi 
            Caption         =   "Instalasi"
            Height          =   375
            Left            =   1320
            TabIndex        =   1
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optJenisPegawai 
            Caption         =   "Jenis Pegawai"
            Height          =   375
            Left            =   2400
            TabIndex        =   2
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkGroup 
         Caption         =   "Group yang dipilih"
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo dcGroup 
         Height          =   390
         Left            =   5640
         TabIndex        =   5
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   688
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Appearance      =   0
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
      TabIndex        =   20
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
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7440
      Picture         =   "frmLaporanSaldoBarang.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLaporanSaldoBarang.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanSaldoBarang.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   4200
      TabIndex        =   10
      Top             =   2880
      Width           =   60
   End
End
Attribute VB_Name = "frmLaporanSaldoBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
            chkNamaBarang.SetFocus
        End If
    End If
End Sub

Private Sub cmdCetak_Click()

    strgroup = "": strNama = ""
    If optHari.Value = True Then
        strCetak = "Hari"
    ElseIf optBulan.Value = True Then
        strCetak = "Bulan"
    Else
        strCetak = ""
    End If

    If optRuangan.Value = True Then
        strgroup = "Ruangan"
    ElseIf optInstalasi.Value = True Then
        strgroup = "Instalasi"
    ElseIf optJenisPegawai.Value = True Then
        strgroup = "JenisPegawai"
    ElseIf optJabatan.Value = True Then
        strgroup = "Jabatan"
    Else
        strgroup = vbNullString
    End If

    If chkGroup.Value = vbChecked Then
        strIsiGroup = dcGroup.Text
    Else
        strIsiGroup = vbNullString
    End If

    'mstrCetak2 = dcTempatBertugas.Text
    strCetak2 = "CetakAbsensi"
    frmCetakAbsensiPegawai.Show
End Sub

Private Sub cmdSpreadSheet_Click()
    strgroup = "": strNama = ""
    
    If optHari.Value = True Then
        strCetak = "Hari"
    ElseIf optBulan.Value = True Then
        strCetak = "Bulan"
    Else
        strCetak = ""
    End If
    
    mdTglAwal = dtpAwal.Value 'TglAwal
    mdTglAkhir = dtpAkhir.Value 'TglAkhir

    If optRuangan.Value = True Then
        strgroup = "Ruangan"
    ElseIf optInstalasi.Value = True Then
        strgroup = "Instalasi"
    ElseIf optJenisPegawai.Value = True Then
        strgroup = "JenisPegawai"
    ElseIf optJabatan.Value = True Then
        strgroup = "Jabatan"
    Else
        strgroup = vbNullString
    End If

    If chkGroup.Value = vbChecked Then
        strIsiGroup = dcGroup.Text
    Else
        strIsiGroup = vbNullString
    End If
       
    frmCetakLapSaldoBarang.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcGroup_GotFocus()
On Error GoTo errload
Dim tempKode As String
    
    tempKode = dcGroup.BoundText
    If optRuangan.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan ORDER BY NamaRuangan")
    ElseIf optInstalasi.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdInstalasi, NamaInstalasi FROM Instalasi ORDER BY NamaInstalasi")
    ElseIf optJenisPegawai.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdJenisPegawai, JenisPegawai FROM JenisPegawai ORDER BY JenisPegawai")
    ElseIf optJabatan.Value = True Then
        Call msubDcSource(dcGroup, rs, "SELECT KdJabatan, NamaJabatan FROM Jabatan ORDER BY NamaJabatan")
    Else
        Exit Sub
    End If
    dcGroup.BoundText = tempKode

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub dcGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcGroup.Text)) = 0 Then dcGroup.SetFocus: Exit Sub
        'If dcGroup.MatchedWithList = True Then chkNamaBarang.SetFocus: Exit Sub
        If optRuangan.Value = True Then
            strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE (NamaRuangan LIKE '%" & dcGroup.Text & "%')"
        ElseIf optInstalasi.Value = True Then
            strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi WHERE NamaInstalasi LIKE '%" & dcGroup.Text & "%')"
        ElseIf optJenisPegawai.Value = True Then
            strSQL = "SELECT KdJenisPegawai, JenisPegawai FROM JenisPegawai WHERE JenisPegawai LIKE '%" & dcGroup.Text & "%')"
        ElseIf optJabatan.Value = True Then
            strSQL = "SELECT KdJabatan, NamaJabatan FROM Jabatan WHERE NamaJabatan LIKE '%" & dcGroup.Text & "%')"
        Else
            optRuangan.SetFocus
        End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcGroup.BoundText = rs(0).Value
        dcGroup.Text = rs(1).Value
    End If
End Sub

Private Sub dcNamaBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcNamaBarang.Text)) = 0 Then dcNamaBarang.SetFocus: Exit Sub
        If dcNamaBarang.MatchedWithList = True Then cmdSpreadSheet.SetFocus: Exit Sub
        Call msubRecFO(rs, "SELECT KdBarang, NamaBarang FROM MasterBarang WHERE NamaBarang LIKE '%" & dcNamaBarang.Text & "%'")
        If rs.EOF = True Then Exit Sub
        dcNamaBarang.BoundText = rs(0).Value
        dcNamaBarang.Text = rs(1).Value
    End If
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optPabrik.SetFocus
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

Private Sub optInstalasi_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optInstalasi.Caption
End Sub

Private Sub optHari_Click()
    dtpAwal.CustomFormat = "dd MMMM yyyy"
    dtpAkhir.CustomFormat = "dd MMMM yyyy"
End Sub

Private Sub optHari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optJenisPegawai_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optJenisPegawai.Caption
End Sub

Private Sub optJenisBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optJabatan_Click()
    dcGroup.BoundText = ""
    chkGroup.Caption = optJabatan.Caption
End Sub

Private Sub optJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkGroup.SetFocus
End Sub

Private Sub optTahun_Click()
    dtpAwal.CustomFormat = "yyyy"
    dtpAkhir.CustomFormat = "yyyy"
End Sub

Private Sub optTahun_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

Private Sub optTotal_Click()
    Call optHari_Click
End Sub
