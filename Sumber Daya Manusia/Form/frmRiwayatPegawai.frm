VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRiwayatPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Pegawai"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRiwayatPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   15930
   Begin VB.Frame Frame2 
      Caption         =   "Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   55
      Top             =   1080
      Width           =   15855
      Begin VB.TextBox txtJenisPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7440
         TabIndex        =   60
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtIdPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         MaxLength       =   10
         TabIndex        =   59
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtNamaPegawai 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   58
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   57
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtJabatan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10320
         TabIndex        =   56
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pegawai"
         Height          =   210
         Left            =   7440
         TabIndex        =   65
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. ID "
         Height          =   210
         Left            =   720
         TabIndex        =   64
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama "
         Height          =   210
         Left            =   2520
         TabIndex        =   63
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5880
         TabIndex        =   62
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan"
         Height          =   210
         Left            =   10320
         TabIndex        =   61
         Top             =   240
         Width           =   630
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   0
      Left            =   0
      TabIndex        =   21
      Top             =   8790
      Width           =   15930
      _ExtentX        =   28099
      _ExtentY        =   0
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   28046
            Text            =   "F5 - Refresh"
            TextSave        =   "F5 - Refresh"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Riwayat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   15855
      Begin VB.CommandButton cmdTutup 
         Caption         =   ">>  TUTU&P"
         Height          =   495
         Left            =   13440
         TabIndex        =   16
         Top             =   6120
         Width           =   2055
      End
      Begin TabDlg.SSTab sstTP 
         Height          =   5655
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   9975
         _Version        =   393216
         Tabs            =   16
         TabsPerRow      =   16
         TabHeight       =   1587
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Riwayat Pendidikan Formal"
         TabPicture(0)   =   "frmRiwayatPegawai.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "dgRiwayatPendidikanFormal"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdTambahRwtPendidikanFormal"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdHapusDataRwtPendidikanFormal"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtTindakanTotal"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Riwayat Pendidikan Non Formal"
         TabPicture(1)   =   "frmRiwayatPegawai.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtAlkesTotal"
         Tab(1).Control(1)=   "cmdHapusNonFormal"
         Tab(1).Control(2)=   "cmdTambahNonFormal"
         Tab(1).Control(3)=   "dgRiwayatPendidikanNonFormal"
         Tab(1).Control(4)=   "Label2"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Riwayat Organisasi"
         TabPicture(2)   =   "frmRiwayatPegawai.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdTambahOrganisasi"
         Tab(2).Control(1)=   "cmdHapusOrganisasi"
         Tab(2).Control(2)=   "dgRiwayatOrganisasi"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Riwayat Perjalanan Dinas"
         TabPicture(3)   =   "frmRiwayatPegawai.frx":0D1E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdcetaksurat"
         Tab(3).Control(1)=   "cmdTambahPerjalananDinas"
         Tab(3).Control(2)=   "cmdHapusPerjalananDinas"
         Tab(3).Control(3)=   "dgRiwayatPerjalananDinas"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "Riwayat Hukuman"
         TabPicture(4)   =   "frmRiwayatPegawai.frx":0D3A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdHapusHukuman"
         Tab(4).Control(1)=   "cmtambahHukuman"
         Tab(4).Control(2)=   "dgRiwayatHukuman"
         Tab(4).ControlCount=   3
         TabCaption(5)   =   "Riwayat Pangkat Golongan"
         TabPicture(5)   =   "frmRiwayatPegawai.frx":0D56
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "cmdHapus"
         Tab(5).Control(1)=   "cmdTambah"
         Tab(5).Control(2)=   "dgRiwayatPangkat"
         Tab(5).ControlCount=   3
         TabCaption(6)   =   "Riwayat Extra Pelatihan"
         TabPicture(6)   =   "frmRiwayatPegawai.frx":0D72
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "cmdTambahExtraPelatihan"
         Tab(6).Control(1)=   "cmdHapusExtraPelatihan"
         Tab(6).Control(2)=   "dgRiwayatPelatihanExtra"
         Tab(6).ControlCount=   3
         TabCaption(7)   =   "Riwayat Prestasi"
         TabPicture(7)   =   "frmRiwayatPegawai.frx":0D8E
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "cmdTambahPrestasi"
         Tab(7).Control(1)=   "cmdHapusPrestasi"
         Tab(7).Control(2)=   "dgRiwayatPrestasi"
         Tab(7).ControlCount=   3
         TabCaption(8)   =   "Riwayat Pekerjaan"
         TabPicture(8)   =   "frmRiwayatPegawai.frx":0DAA
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "dgRiwayatPekerjaan"
         Tab(8).Control(0).Enabled=   0   'False
         Tab(8).Control(1)=   "chkRP"
         Tab(8).Control(1).Enabled=   0   'False
         Tab(8).Control(2)=   "cmdHapusPekerjaan"
         Tab(8).Control(2).Enabled=   0   'False
         Tab(8).Control(3)=   "cmdTambahPekerjaan"
         Tab(8).Control(3).Enabled=   0   'False
         Tab(8).ControlCount=   4
         TabCaption(9)   =   "Riwayat Gaji"
         TabPicture(9)   =   "frmRiwayatPegawai.frx":0DC6
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "cmdTambahGaji"
         Tab(9).Control(1)=   "cmdHapusGaji"
         Tab(9).Control(2)=   "cmdCetakSuratBerkala"
         Tab(9).Control(3)=   "dgRiwayatGaji"
         Tab(9).ControlCount=   4
         TabCaption(10)  =   "Riwayat Jabatan/Ruang/Mutasi"
         TabPicture(10)  =   "frmRiwayatPegawai.frx":0DE2
         Tab(10).ControlEnabled=   0   'False
         Tab(10).Control(0)=   "cmdHapusRiwayatTempatBertugas"
         Tab(10).Control(1)=   "cmdTambahRiwayatTempatBertugas"
         Tab(10).Control(2)=   "dgRiwayatTempatBertugas"
         Tab(10).ControlCount=   3
         TabCaption(11)  =   "Riwayat Keluarga"
         TabPicture(11)  =   "frmRiwayatPegawai.frx":0DFE
         Tab(11).ControlEnabled=   0   'False
         Tab(11).Control(0)=   "cmdHapusRiwayatKeluarga"
         Tab(11).Control(1)=   "cmdRiwayatKeluarga"
         Tab(11).Control(2)=   "dgRiwayatKeluarga"
         Tab(11).ControlCount=   3
         TabCaption(12)  =   "Riwayat Status"
         TabPicture(12)  =   "frmRiwayatPegawai.frx":0E1A
         Tab(12).ControlEnabled=   0   'False
         Tab(12).Control(0)=   "cmdtambahriwayatstatus"
         Tab(12).Control(1)=   "cmdhapusriwayatstatus"
         Tab(12).Control(2)=   "cmdcetaksuratcuti"
         Tab(12).Control(3)=   "dgRiwayatSta"
         Tab(12).ControlCount=   4
         TabCaption(13)  =   "Riwayat Tugas Belajar"
         TabPicture(13)  =   "frmRiwayatPegawai.frx":0E36
         Tab(13).ControlEnabled=   0   'False
         Tab(13).Control(0)=   "cmdCetakTugas"
         Tab(13).Control(1)=   "cmdHapusTugas"
         Tab(13).Control(2)=   "cmdTambahTugas"
         Tab(13).Control(3)=   "dgTugas"
         Tab(13).ControlCount=   4
         TabCaption(14)  =   "Riwayat Potongan Gaji"
         TabPicture(14)  =   "frmRiwayatPegawai.frx":0E52
         Tab(14).ControlEnabled=   0   'False
         Tab(14).Control(0)=   "CmdTambahPotonganGaji"
         Tab(14).Control(1)=   "dgRiwayatPotonganGaji"
         Tab(14).ControlCount=   2
         TabCaption(15)  =   "Surat Izin Praktik"
         TabPicture(15)  =   "frmRiwayatPegawai.frx":0E6E
         Tab(15).ControlEnabled=   0   'False
         Tab(15).Control(0)=   "dgRiwayatSIP"
         Tab(15).Control(1)=   "cmdTambahSIP"
         Tab(15).ControlCount=   2
         Begin VB.CommandButton cmdTambahSIP 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -61200
            TabIndex        =   70
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton CmdTambahPotonganGaji 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   68
            Top             =   5160
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakTugas 
            Caption         =   "Cetak Surat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65760
            TabIndex        =   54
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusTugas 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   53
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahTugas 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   52
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdtambahriwayatstatus 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   50
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdhapusriwayatstatus 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   49
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdcetaksuratcuti 
            Caption         =   "Cetak Surat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65760
            TabIndex        =   48
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusRiwayatKeluarga 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   46
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdRiwayatKeluarga 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   45
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusRiwayatTempatBertugas 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   43
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahRiwayatTempatBertugas 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   42
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahGaji 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   40
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusGaji 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   39
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetakSuratBerkala 
            Caption         =   "Cetak Berkala"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65760
            TabIndex        =   38
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPekerjaan 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   36
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusPekerjaan 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   35
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPrestasi 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   33
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusPrestasi 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   32
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahExtraPelatihan 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   30
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusExtraPelatihan 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   29
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapus 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   27
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambah 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   26
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusHukuman 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   24
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmtambahHukuman 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   23
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdcetaksurat 
            Caption         =   "Cetak Surat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65760
            TabIndex        =   22
            Top             =   5130
            Width           =   1575
         End
         Begin VB.TextBox txtTindakanTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3000
            TabIndex        =   2
            Top             =   8220
            Width           =   2415
         End
         Begin VB.TextBox txtAlkesTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   -71640
            TabIndex        =   6
            Top             =   8220
            Width           =   2415
         End
         Begin VB.CommandButton cmdHapusDataRwtPendidikanFormal 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10920
            TabIndex        =   3
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahRwtPendidikanFormal 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12600
            TabIndex        =   4
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusNonFormal 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   7
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahNonFormal 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   8
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CheckBox chkRP 
            Caption         =   "Tampilkan Semua"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -74760
            TabIndex        =   15
            Top             =   8100
            Width           =   1815
         End
         Begin VB.CommandButton cmdTambahOrganisasi 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   11
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusOrganisasi 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   10
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdTambahPerjalananDinas 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -62400
            TabIndex        =   14
            Top             =   5130
            Width           =   1575
         End
         Begin VB.CommandButton cmdHapusPerjalananDinas 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64080
            TabIndex        =   13
            Top             =   5130
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid dgRiwayatPendidikanFormal 
            Height          =   3855
            Left            =   120
            TabIndex        =   1
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatPendidikanNonFormal 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   5
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatOrganisasi 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   9
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatPerjalananDinas 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   12
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatHukuman 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   25
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatPangkat 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   28
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatPelatihanExtra 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   31
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatPrestasi 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   34
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatPekerjaan 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   37
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatGaji 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   41
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatTempatBertugas 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   44
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatKeluarga 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   47
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatSta 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   51
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgTugas 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   66
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatPotonganGaji 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   67
            Top             =   1080
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgRiwayatSIP 
            Height          =   3855
            Left            =   -74760
            TabIndex        =   69
            Top             =   1080
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pelayanan Tindakan"
            Height          =   210
            Left            =   240
            TabIndex        =   19
            Top             =   8280
            Width           =   2550
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pemakaian Obat && Alkes"
            Height          =   210
            Left            =   -74760
            TabIndex        =   18
            Top             =   8280
            Width           =   2925
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   1800
      _cx             =   4197479
      _cy             =   4196024
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatPegawai.frx":0E8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   14040
      Picture         =   "frmRiwayatPegawai.frx":384B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPegawai.frx":45D3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmRiwayatPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer

Private Sub cmdcetaksurat_Click()
    On Error GoTo hell
    'If txtIdPegawai.Text = "" Then Exit Sub
    If dgRiwayatPerjalananDinas.ApproxCount = 0 Then Exit Sub
    strNoUrut = dgRiwayatPerjalananDinas.Columns("No. Urut")
    frmPilihPenandatangan.Show
    frmRiwayatPegawai.Enabled = False

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdCetakSuratBerkala_Click()
On Error Resume Next
    Dim pesan As VbMsgBoxResult
    If dgRiwayatGaji.ApproxCount = 0 Then Exit Sub
    'strSQL = "select * from V_CetakGajiBerkala where idpegawai='" & dgRiwayatGaji.Columns("IdPegawai") & "' AND KdKomponenGaji = '" & dgRiwayatGaji.Columns("KdKomponenGaji").Value & "' "
    strSQL = "select * from V_CetakGajiBerkala where idpegawai='" & dgRiwayatGaji.Columns("IdPegawai") & "' AND NoUrut='" & dgRiwayatGaji.Columns("No. Urut") & "'"

    Call msubRecFO(rs, strSQL)
    strNoUrut = rs.Fields("NoUrut").Value
    strIDPegawai = rs.Fields("IdPegawai").Value

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"
    frmCetakSuratBerkala.KdKomponenGaji = rs.Fields("KdKomponenGaji").Value
    frmCetakSuratBerkala.Show
Exit Sub
    Call msubPesanError
End Sub

Private Sub cmdcetaksuratcuti_Click()
    On Error GoTo hell
    
    If dgRiwayatSta.ApproxCount = 0 Then Exit Sub
    strSQL = "select * from V_CetakSuratCuti where idpegawai='" & mstrIdPegawai & "' AND NoRiwayat='" & dgRiwayatSta.Columns("NoRiwayat") & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        MsgBox "Tidak ada data, silahkan lengkapi kelengkapan data pegawai bersangkutan", vbExclamation, "Validasi"
    Exit Sub
    End If
'    If rs.Fields("KdStatus").Value = "02" Or rs.Fields("KdStatus").Value = "07" _
'        Or rs.Fields("KdStatus").Value = "08" Or rs.Fields("KdStatus").Value = "09" Then
''        Or rs.Fields("KdStatus").Value = "10" Or rs.Fields("KdStatus").Value = "01" Then

    If rs.Fields("KdStatus").Value = "02" Then 'cuti
        frmCetakSuratKeteranganCutiTahunan.Show
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdCetakTugas_Click()
On Error Resume Next
    Dim pesan As VbMsgBoxResult
    If dgTugas.ApproxCount = 0 Then Exit Sub
    strSQL = "select * from V_RiwayatTugas where idpegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgTugas.Columns("No. Urut") & "'"
    Call msubRecFO(rs, strSQL)

    If rs.EOF = True Then
        MsgBox "Tidak ada data  ", vbCritical, "Validasi"
        Exit Sub
    End If

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    frmCetakSuratTugas.Show
Exit Sub

    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    Dim tempNoUrutX As String
    Dim tempKdPangkatX As String
    If dgRiwayatPangkat.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Pangkat ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "Delete From RiwayatPangkat WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatPangkat.Columns("No. Urut").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatPangkat
    strSQL = "select max(NoUrut) from V_BantuUpdatePangkatToCurrent Where IdPegawai = '" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    tempNoUrutX = rs(0).Value
    strSQL = "select KdPangkat from Pangkat where NoUrut='" & tempNoUrutX & "'"
    Call msubRecFO(rs, strSQL)
    tempKdPangkatX = rs(0).Value
    strSQL = "update DataCurrentPegawai set KdPangkat='" & tempKdPangkatX & "' where IdPegawai='" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)

    Exit Sub
errHapus:
End Sub

Private Sub cmdHapusDataRwtPendidikanFormal_Click()
    On Error GoTo errHapus
    If dgRiwayatPendidikanFormal.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Pendidikan Formal ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPendidikanFormal WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatPendidikanFormal.Columns("No. Urut").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatPendidikanFormal
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusExtraPelatihan_Click()
    On Error GoTo errHapus
    If dgRiwayatPelatihanExtra.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Extra Pelatihan ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatExtraPelatihan WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatPelatihanExtra.Columns("No. Urut").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatExtraPelatihan
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusGaji_Click()
    On Error GoTo errHapus
    If dgRiwayatGaji.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Gaji ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM DetailRiwayatGaji WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatGaji.Columns("No. Urut").Value & "' AND KdKomponenGaji = '" & dgRiwayatGaji.Columns("KdKomponenGaji").Value & "'"
    strSQL = "DELETE FROM RiwayatGaji WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatGaji.Columns("No. Urut").Value & "' "
    dbConn.Execute strSQL
    Call subLoadRiwayatGaji
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusHukuman_Click()
    On Error GoTo errHapus
    If dgRiwayatHukuman.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Hukuman ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatHukuman WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatHukuman.Columns("No. Urut").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatHukuman
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusNonFormal_Click()
    On Error GoTo errHapus
    If dgRiwayatPendidikanNonFormal.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Pendidikan Non Formal ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPendidikanNonFormal WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatPendidikanNonFormal.Columns("No. Urut").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatPendidikanNonFormal
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusOrganisasi_Click()
    On Error GoTo errHapus
    If dgRiwayatOrganisasi.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Organisasi ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatOrganisasi WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatOrganisasi.Columns("No. Urut").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatOrganisasi
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusPekerjaan_Click()
    On Error GoTo errHapus
    If dgRiwayatPekerjaan.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Pekerjaan ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPekerjaan WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatPekerjaan.Columns("No. Urut").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatPekerjaan
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusPerjalananDinas_Click()
    On Error GoTo errHapus
    If dgRiwayatPerjalananDinas.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Perjalanan Dinas ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPerjalananDinas WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatPerjalananDinas.Columns("No. Urut").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatPerjalananDinas
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusPrestasi_Click()
    On Error GoTo errHapus
    If dgRiwayatPrestasi.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Prestasi ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPrestasi WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & dgRiwayatPrestasi.Columns("No. Urut").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatPrestasi
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusRiwayatKeluarga_Click()
    On Error GoTo errHapus
    If dgRiwayatKeluarga.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Keluarga ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM KeluargaPegawai WHERE IdPegawai='" & mstrIdPegawai & "' AND KdHubungan = '" & dgRiwayatKeluarga.Columns("KdHubungan").Value & "' AND NoUrut = '" & dgRiwayatKeluarga.Columns("No. Urut").Value & "' "
    dbConn.Execute strSQL
    Call subLoadRiwayatKeluargaPegawai
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdhapusriwayatstatus_Click()
    On Error GoTo errHapus
    If dgRiwayatSta.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Status ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatStatusPegawai WHERE IdPegawai='" & mstrIdPegawai & "' AND NoRiwayat='" & dgRiwayatSta.Columns("NoRiwayat").Value & "'"
    dbConn.Execute strSQL
    strSQL = "DELETE FROM Riwayat WHERE NoRiwayat='" & dgRiwayatSta.Columns("NoRiwayat").Value & "'"
    dbConn.Execute strSQL
    Call subLoadRiwayatStatus

    Dim tempNoUrutX As String
    Dim tempKdStatusX As String
    Dim tempTglAwal As Date
    strSQL = "select max(NoRiwayat) from RiwayatStatusPegawai Where IdPegawai = '" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    tempNoUrutX = rs(0).Value
    strSQL = "select KdStatus, TglAwal from RiwayatStatusPegawai where NoRiwayat='" & tempNoUrutX & "' and IdPegawai = '" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    tempKdStatusX = rs(0).Value
    tempTglAwal = rs(1).Value
    strSQL = "update DataCurrentPegawai set KdStatus='" & tempKdStatusX & "' where IdPegawai='" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    If tempKdStatusX = "01" Then
        strSQL = "update DataPegawai set TglKeluar=Null where IdPegawai='" & mstrIdPegawai & "'"
        Call msubRecFO(rs, strSQL)
    Else
        strSQL = "update DataPegawai set TglKeluar='" & Format(tempTglAwal, "yyyy/MM/dd") & "' where IdPegawai='" & mstrIdPegawai & "'"
        Call msubRecFO(rs, strSQL)
    End If

    Exit Sub
errHapus:
End Sub

Private Sub cmdHapusRiwayatTempatBertugas_Click()
    On Error GoTo errHapus
    If dgRiwayatTempatBertugas.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Tempat Bertugas ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatTempatBertugas WHERE NoRiwayat='" & dgRiwayatTempatBertugas.Columns("NoRiwayat").Value & "' and IdPegawai='" & mstrIdPegawai & "' "
    dbConn.Execute strSQL
    strSQL = "DELETE FROM Riwayat WHERE NoRiwayat='" & dgRiwayatTempatBertugas.Columns("NoRiwayat").Value & "' "
    dbConn.Execute strSQL
    Call subLoadRiwayatTempatBertugas
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdHapusTugas_Click()
    On Error GoTo errHapus
    If dgTugas.ApproxCount = 0 Then Exit Sub
    If MsgBox("Hapus Riwayat Tugas ", vbInformation + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatTugas WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut = '" & dgTugas.Columns("No. Urut").Value & "' "
    dbConn.Execute strSQL
    Call subLoadRiwayatTugas
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdRiwayatKeluarga_Click()
    frmRiwayatKeluargaPegawai.Show
    Me.Enabled = False
End Sub

Private Sub cmdTambah_Click()
    frmRiwayatPangkat.Show
    Me.Enabled = False
End Sub

Private Sub cmdTambahExtraPelatihan_Click()
    frmRiwayatExtraPelatihan.Show
    Me.Enabled = False
End Sub

Private Sub cmdTambahGaji_Click()
    frmRiwayatGaji.Show
    Me.Enabled = False
End Sub

Private Sub cmdTambahJabatan_Click()
    frmRiwayatJabatan.Show
    Me.Enabled = False
End Sub

Private Sub cmdTambahKawin_Click()
    frmRiwayatPerkawinan.Show
    Me.Enabled = False
End Sub

Private Sub cmdTambahMutasiKepegwaian_Click()
    frmMutasiPegawai.Show
End Sub

Private Sub cmdTambahNonFormal_Click()
    On Error GoTo errLoad
    frmRiwayatPendidikanNonFormal.Show
    Me.Enabled = False
    Exit Sub
errLoad:
End Sub

Private Sub cmdTambahOrganisasi_Click()
    frmRiwayatOrganisasi.Show
    Me.Enabled = False
End Sub

Private Sub cmdTambahPekerjaan_Click()
    frmRiwayatPekerjaan.Show
    Me.Enabled = False

End Sub

Private Sub cmdTambahPerjalananDinas_Click()
    frmRiwayatPerjalananDinas.Show
    Me.Enabled = False
End Sub

Private Sub CmdTambahPotonganGaji_Click()
    frmRiwayatPotonganGaji.Show
End Sub

Private Sub cmdTambahPrestasi_Click()
    frmRiwayatPrestasi.Show
    Me.Enabled = False
End Sub

Private Sub cmdtambahriwayatstatus_Click()
    frmRiwayatStatusPegawai.Show
    Me.Enabled = False
End Sub

Private Sub cmdTambahRiwayatTempatBertugas_Click()
    If txtIdPegawai = "" Then Exit Sub
    mstrIdPegawai = txtIdPegawai.Text
    frmRiwayatTempatBertugas.Show
    Me.Enabled = False
End Sub

Private Sub cmdTambahRwtPendidikanFormal_Click()
    On Error GoTo errLoad
    frmRiwayatPendidikanFormal.Show
    Me.Enabled = False
    Exit Sub
errLoad:
End Sub

Private Sub cmdTambahSIP_Click()
    frmSuratSIP.Show
End Sub

Private Sub cmdTambahTugas_Click()
    frmRiwayatTugas.Show
    Me.Enabled = False
End Sub

Private Sub cmdTutup_Click()
    Unload Me
    frmDataPegawaiNew.Enabled = True
End Sub

Private Sub cmtambahHukuman_Click()
    frmRiwayatHukuman.Show
    Me.Enabled = False
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad

    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKey1
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 0
        Case vbKey2
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 1
        Case vbKey3
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 2
        Case vbKey4
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 3
        Case vbKey5
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 4
        Case vbKey6
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 5
        Case vbKey7
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 6
        Case vbKey8
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 7
        Case vbKey9
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 8
        Case vbKey0
            If strCtrlKey = 4 Then sstTP.SetFocus: sstTP.Tab = 9
        Case vbKeyF5
            Call subLoadRiwayatPendidikanFormal
            Call subLoadRiwayatPendidikanNonFormal
            Call subLoadRiwayatOrganisasi
            Call subLoadRiwayatPerjalananDinas
            Call subLoadRiwayatHukuman
            Call subLoadRiwayatPangkat
            Call subLoadRiwayatExtraPelatihan
            Call subLoadRiwayatPrestasi
            Call subLoadRiwayatPekerjaan
            Call subLoadRiwayatGaji
            Call subLoadRiwayatTempatBertugas
            Call subLoadRiwayatKeluargaPegawai
            Call subLoadRiwayatStatus
            Call subLoadRiwayatTugas
            Call subLoadRiwayatSIP
            Call subLoadRiwayatMutasiKepegawaian
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    txtIdPegawai.Text = frmDataPegawaiNew.txtIdPegawai.Text
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Call subLoadRiwayatPendidikanFormal
'    Call subLoadRiwayatPendidikanNonFormal
'    Call subLoadRiwayatOrganisasi
'    Call subLoadRiwayatPerjalananDinas
'    Call subLoadRiwayatHukuman
'    Call subLoadRiwayatPangkat
'    Call subLoadRiwayatExtraPelatihan
'    Call subLoadRiwayatPrestasi
'    Call subLoadRiwayatPekerjaan
'    Call subLoadRiwayatGaji
'    Call subLoadRiwayatTempatBertugas
'    Call subLoadRiwayatKeluargaPegawai
'    Call subLoadRiwayatStatus
'    Call subLoadRiwayatTugas
'    Call subLoadRiwayatPotongan
'    Call subLoadRiwayatSIP
'    Call subLoadRiwayatMutasiKepegawaian

    sstTP.Tab = 0
End Sub

'Untuk meload riwayat pendidikan formal pegawai
Public Sub subLoadRiwayatPendidikanFormal()
    On Error GoTo hell
    strSQL = "SELECT [No. Urut], Pendidikan, [Nama Sekolah], Jurusan, [Tgl. Masuk], [Tgl. Lulus], IPK, Kelulusan, [No. Ijazah], [Tgl. Ijazah], [TTD Ijazah], " & _
    " [Alamat Sekolah] , [Pimpinan Sekolah], Keterangan, [Nama User]" & _
    " FROM v_RiwayatPendidikanFormal where [ID Peg] = '" & mstrIdPegawai & "'  "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatPendidikanFormal.DataSource = rs
    With dgRiwayatPendidikanFormal
        .Columns("No. Urut").Width = 1000
        .Columns("No. Urut").Alignment = vbCenter
        .Columns("Pendidikan").Width = 1500
        .Columns("Nama Sekolah").Width = 2000
        .Columns("Jurusan").Width = 1000
        .Columns("Tgl. Masuk").Width = 1100
        .Columns("Tgl. Lulus").Width = 1100
        .Columns("IPK").Width = 700
        .Columns("Kelulusan").Width = 1700
        .Columns("No. Ijazah").Width = 2000
        .Columns("Tgl. Ijazah").Width = 1500
        .Columns("TTD Ijazah").Width = 2000
        .Columns("Alamat Sekolah").Width = 2500
        .Columns("Pimpinan Sekolah").Width = 2500
        .Columns("Keterangan").Width = 2500
        .Columns("Nama User").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat pendidikan non formal pegawai
Public Sub subLoadRiwayatPendidikanNonFormal()
    On Error GoTo hell
    strSQL = "SELECT dbo.RiwayatPendidikanNonFormal.NoUrut, dbo.RiwayatPendidikanNonFormal.NamaPendidikan, dbo.RiwayatPendidikanNonFormal.LamaPendidikan, " & _
    " dbo.RiwayatPendidikanNonFormal.TglMulai, dbo.RiwayatPendidikanNonFormal.TglLulus, dbo.RiwayatPendidikanNonFormal.NoSertifikat," & _
    " dbo.RiwayatPendidikanNonFormal.TglSertifikat, dbo.RiwayatPendidikanNonFormal.TandaTanganSertifikat," & _
    " dbo.RiwayatPendidikanNonFormal.InstansiPendidikan, dbo.RiwayatPendidikanNonFormal.AlamatPendidikan, dbo.RiwayatPendidikanNonFormal.PimpinanPendidikan," & _
    " dbo.RiwayatPendidikanNonFormal.Keterangan, dbo.DataPegawai.NamaLengkap AS NamaUser" & _
    " FROM dbo.RiwayatPendidikanNonFormal INNER JOIN " & _
    " dbo.DataPegawai ON dbo.RiwayatPendidikanNonFormal.IdUser = dbo.DataPegawai.IdPegawai where RiwayatPendidikanNonFormal.IdPegawai = '" & mstrIdPegawai & "'  "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatPendidikanNonFormal.DataSource = rs
    With dgRiwayatPendidikanNonFormal
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Alignment = vbCenter
        .Columns("NoUrut").Caption = "No. Urut"

        .Columns("NamaPendidikan").Caption = "Nama Pendidikan"
        .Columns("LamaPendidikan").Width = 1500
        .Columns("LamaPendidikan").Alignment = vbCenter
        .Columns("LamaPendidikan").Caption = "Lama Pendidikan"
        .Columns("TglMulai").Width = 1200
        .Columns("TglMulai").Caption = "Tgl. Mulai"
        .Columns("TglLulus").Width = 1200
        .Columns("TglLulus").Caption = "Tgl. Lulus"
        .Columns("NoSertifikat").Width = 1200
        .Columns("NoSertifikat").Caption = "No. Sertifikat"
        .Columns("TglSertifikat").Width = 1200
        .Columns("TglSertifikat").Caption = "Tgl. Sertifikat"
        .Columns("TandaTanganSertifikat").Width = 2000
        .Columns("TandaTanganSertifikat").Caption = "TTD Sertifikat"
        .Columns("InstansiPendidikan").Width = 3000
        .Columns("InstansiPendidikan").Caption = "Instansi Pendidikan"
        .Columns("AlamatPendidikan").Width = 2000
        .Columns("AlamatPendidikan").Caption = "Alamat Pendidikan"
        .Columns("PimpinanPendidikan").Width = 2000
        .Columns("PimpinanPendidikan").Caption = "Pimpinan Pendidikan"
        .Columns("Keterangan").Width = 2000
        .Columns("NamaUser").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat organisasi
Public Sub subLoadRiwayatOrganisasi()
    On Error GoTo hell
    strSQL = "SELECT dbo.RiwayatOrganisasi.NoUrut, dbo.RiwayatOrganisasi.NamaOrganisasi, dbo.RiwayatOrganisasi.Jabatan, dbo.RiwayatOrganisasi.TglMasuk, " & _
    " dbo.RiwayatOrganisasi.TglAkhir, dbo.RiwayatOrganisasi.AlamatOrganisasi, dbo.RiwayatOrganisasi.PimpinanOrganisasi," & _
    " dbo.RiwayatOrganisasi.Keterangan, dbo.DataPegawai.NamaLengkap AS NamaUser" & _
    " FROM dbo.RiwayatOrganisasi INNER JOIN " & _
    " dbo.DataPegawai ON dbo.RiwayatOrganisasi.IdUser = dbo.DataPegawai.IdPegawai where RiwayatOrganisasi.IdPegawai = '" & mstrIdPegawai & "'  "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatOrganisasi.DataSource = rs
    With dgRiwayatOrganisasi
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Alignment = vbCenter
        .Columns("NoUrut").Caption = "No. Urut"

        .Columns("NamaOrganisasi").Width = 2000
        .Columns("NamaOrganisasi").Caption = "Nama Organisasi"
        .Columns("Jabatan").Width = 2000
        .Columns("TglMasuk").Width = 1200
        .Columns("TglMasuk").Caption = "Tgl. Masuk"
        .Columns("TglAkhir").Width = 1200
        .Columns("TglAkhir").Caption = "Tgl. AKhir"
        .Columns("AlamatOrganisasi").Width = 3000
        .Columns("AlamatOrganisasi").Caption = "Alamat Organisasi"
        .Columns("PimpinanOrganisasi").Width = 2100
        .Columns("PimpinanOrganisasi").Caption = "Pimpinan Organisasi"
        .Columns("Keterangan").Width = 3100
        .Columns("NamaUser").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat tempat bertugas
Public Sub subLoadRiwayatTempatBertugas()
    On Error GoTo hell
    strSQL = "SELECT * FROM V_RiwayatTempatbertugas where IdPegawai = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatTempatBertugas.DataSource = rs
    With dgRiwayatTempatBertugas

        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i

        .Columns(3).Width = 2000 'ruangan
        .Columns(4).Width = 2000 'satuan kerja
        .Columns(5).Width = 2000 'jabatan
        .Columns(6).Width = 1000 'eselon
        .Columns(7).Width = 1000 'tglsk
        .Columns(8).Width = 1500 'nosk
        .Columns(9).Width = 1500 'tandatanGan sk
        .Columns(10).Width = 1000 'tglmulai berlaku
        .Columns(11).Width = 1000 'tglakhirberlaku
        .Columns(12).Width = 2000 'keteranga

    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat keluarga pegawai
Public Sub subLoadRiwayatKeluargaPegawai()
    On Error GoTo hell
    strSQL = "SELECT *" & _
    " FROM V_KeluargaPegawai where IdPegawai = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatKeluarga.DataSource = rs
    With dgRiwayatKeluarga
        .Columns("IdPegawai").Width = 0
        .Columns("KdHubungan").Width = 0
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Caption = "No. Urut"
        .Columns("NamaHubungan").Width = 2000
        .Columns("NamaHubungan").Caption = "Hubungan"
        .Columns("NamaKeluarga").Width = 3000
        '.Columns("NamaKeluarga").Caption = "Nama Keluarga"
        .Columns("JenisKelamin").Width = 500
        .Columns("JenisKelamin").Alignment = vbCenter
        .Columns("JenisKelamin").Caption = "JK"
        .Columns("TglLahir").Width = 1200
        .Columns("TgLLahir").Caption = "Tgl. Lahir"
        .Columns("Pekerjaan").Width = 2500
        .Columns("Pendidikan").Width = 2000
        .Columns("Keterangan").Width = 3600
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat perjalanan dinas
Public Sub subLoadRiwayatPerjalananDinas()
    On Error GoTo hell
    strSQL = "SELECT dbo.RiwayatPerjalananDinas.NoUrut, dbo.RiwayatPerjalananDinas.KotaTujuan, dbo.RiwayatPerjalananDinas.NegaraTujuan, " & _
    " dbo.RiwayatPerjalananDinas.TujuanKunjungan, dbo.RiwayatPerjalananDinas.TglPergi, dbo.RiwayatPerjalananDinas.TglPulang," & _
    " dbo.RiwayatPerjalananDinas.PenyandangDana, dbo.RiwayatPerjalananDinas.Keterangan, dbo.DataPegawai.NamaLengkap AS NamaUser" & _
    " FROM dbo.RiwayatPerjalananDinas INNER JOIN" & _
    " dbo.DataPegawai ON dbo.RiwayatPerjalananDinas.IdUser = dbo.DataPegawai.IdPegawai where RiwayatPerjalananDinas.IdPegawai = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatPerjalananDinas.DataSource = rs
    With dgRiwayatPerjalananDinas
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Alignment = vbCenter
        .Columns("NoUrut").Caption = "No. Urut"

        .Columns("KotaTujuan").Width = 2300
        .Columns("KotaTujuan").Caption = "Kota Tujuan"
        .Columns("NegaraTujuan").Width = 2300
        .Columns("NegaraTujuan").Caption = "Negara Tujuan"
        .Columns("TujuanKunjungan").Width = 2000
        .Columns("TujuanKunjungan").Caption = "Tujuan Kunjungan"
        .Columns("TglPergi").Width = 1200
        .Columns("TglPergi").Caption = "Tgl. Pergi"
        .Columns("TglPulang").Width = 1200
        .Columns("TglPulang").Caption = "Tgl. Pulang"
        .Columns("PenyandangDana").Width = 2350
        .Columns("PenyandangDana").Caption = "Penyandang Dana"
        .Columns("Keterangan").Width = 2350
        .Columns("NamaUser").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat hukuman
Public Sub subLoadRiwayatHukuman()
    On Error GoTo hell
    strSQL = "SELECT NoUrut, JenisHukuman, NoSK, TglSK, TMT, TandaTanganSK, Keterangan, TglSelesai, NamaUser" & _
    " FROM v_RiwayatHukuman where IdPegawai = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatHukuman.DataSource = rs
    With dgRiwayatHukuman
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Alignment = vbCenter
        .Columns("NoUrut").Caption = "No. Urut"

        .Columns("JenisHukuman").Width = 2000
        .Columns("JenisHukuman").Caption = "Jenis Hukuman"
        .Columns("NoSK").Width = 2000
        .Columns("NoSK").Caption = "No. SK"
        .Columns("TglSK").Width = 1200
        .Columns("TglSK").Caption = "Tgl. SK"
        .Columns("TMT").Width = 1200
        .Columns("TandaTanganSK").Width = 1600
        .Columns("TandaTanganSK").Caption = "TTD SK"
        .Columns("Keterangan").Width = 2200
        .Columns("TglSelesai").Width = 1200
        .Columns("TglSelesai").Caption = "Tgl. Selesai"
        .Columns("NamaUser").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat pangkat
Public Sub subLoadRiwayatPangkat()
    On Error GoTo hell
    strSQL = "SELECT  NoUrut, NamaPangkat, NamaGolongan, NoSK, TglSK, TMT, TandaTanganSK, Keterangan, NamaUser" & _
    " FROM v_RiwayatPangkat where IdPegawai = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatPangkat.DataSource = rs
    With dgRiwayatPangkat
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Alignment = vbCenter
        .Columns("NoUrut").Caption = "No. Urut"

        .Columns("NamaPangkat").Width = 2600
        .Columns("NamaPangkat").Caption = "Pangkat"
        .Columns("NamaGolongan").Width = 1000
        .Columns("NamaGolongan").Caption = "Golongan"
        .Columns("NoSK").Width = 2000
        .Columns("NoSK").Caption = "No. SK"
        .Columns("TglSK").Width = 1200
        .Columns("TglSK").Caption = "Tgl. SK"
        .Columns("TandaTanganSK").Width = 1600
        .Columns("TandaTanganSK").Caption = "TTD SK"
        .Columns("Keterangan").Width = 3000
        .Columns("TMT").Width = 1200
        .Columns("NamaUser").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat extra pelatihan
Public Sub subLoadRiwayatExtraPelatihan()
    On Error GoTo hell
    strSQL = " SELECT     dbo.RiwayatExtraPelatihan.NoUrut, dbo.RiwayatExtraPelatihan.NamaPelatihan, dbo.RiwayatExtraPelatihan.KedudukanPeranan, " & _
    " dbo.RiwayatExtraPelatihan.TglMulai, dbo.RiwayatExtraPelatihan.TglAkhir, dbo.RiwayatExtraPelatihan.lamawaktu, dbo.RiwayatExtraPelatihan.InstansiPenyelenggara," & _
    " dbo.RiwayatExtraPelatihan.AlamatPenyelenggara, dbo.RiwayatExtraPelatihan.PimpinanPenyelenggara, dbo.RiwayatExtraPelatihan.Keterangan," & _
    " dbo.DataPegawai.NamaLengkap AS NamaUser" & _
    " FROM dbo.RiwayatExtraPelatihan INNER JOIN" & _
    " dbo.DataPegawai ON dbo.RiwayatExtraPelatihan.IdUser = dbo.DataPegawai.IdPegawai where RiwayatExtraPelatihan.IdPegawai = '" & mstrIdPegawai & "'  "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatPelatihanExtra.DataSource = rs
    With dgRiwayatPelatihanExtra
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Alignment = vbCenter
        .Columns("NoUrut").Caption = "No. Urut"
        .Columns("NamaPelatihan").Width = 2150
        .Columns("NamaPelatihan").Caption = "Nama Pelatihan"
        .Columns("KedudukanPeranan").Width = 2000
        .Columns("KedudukanPeranan").Caption = "Peranan"
        .Columns("TglMulai").Width = 1500
        .Columns("TglMulai").Caption = "Tgl. Mulai"
        .Columns("TglAkhir").Width = 1500
        .Columns("TglAkhir").Caption = "Tgl. Akhir"
        .Columns("LamaWaktu").Width = 1500
        .Columns("LamaWaktu").Caption = "Lama Waktu"
        .Columns("InstansiPenyelenggara").Width = 1600
        .Columns("InstansiPenyelenggara").Caption = "Penyelenggara"
        .Columns("AlamatPenyelenggara").Width = 2500
        .Columns("AlamatPenyelenggara").Caption = "Lokasi"
        .Columns("PimpinanPenyelenggara").Width = 2500
        .Columns("PimpinanPenyelenggara").Caption = "Pimpinan"
        .Columns("Keterangan").Width = 3500
        .Columns("NamaUser").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat prestasi
Public Sub subLoadRiwayatPrestasi()
    On Error GoTo hell
    strSQL = " SELECT dbo.RiwayatPrestasi.NoUrut, dbo.RiwayatPrestasi.NamaPenghargaan, dbo.RiwayatPrestasi.TglDiperoleh, dbo.RiwayatPrestasi.InstansiPemberi, " & _
    " dbo.RiwayatPrestasi.PimpinanInstansiPemberi, dbo.RiwayatPrestasi.NomorPiagam, dbo.RiwayatPrestasi.Keterangan, dbo.DataPegawai.NamaLengkap AS NamaUser" & _
    " FROM dbo.RiwayatPrestasi INNER JOIN" & _
    " dbo.DataPegawai ON dbo.RiwayatPrestasi.IdUser = dbo.DataPegawai.IdPegawai where RiwayatPrestasi.IdPegawai = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatPrestasi.DataSource = rs
    With dgRiwayatPrestasi
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Alignment = vbCenter
        .Columns("NoUrut").Caption = "No. Urut"

        .Columns("NamaPenghargaan").Width = 2500
        .Columns("NamaPenghargaan").Caption = "Penghargaan"
        .Columns("TglDiperoleh").Width = 1200
        .Columns("TglDiperoleh").Caption = "Tgl. Diperoleh"
        .Columns("InstansiPemberi").Width = 2000
        .Columns("InstansiPemberi").Caption = "Instansi Pemberi"
        .Columns("PimpinanInstansiPemberi").Width = 2400
        .Columns("PimpinanInstansiPemberi").Caption = "Pimpinan"
        .Columns("NomorPiagam").Width = 1200
        .Columns("Keterangan").Width = 4350
        .Columns("NamaUser").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat pekerjaan
Public Sub subLoadRiwayatPekerjaan()
    On Error GoTo hell
    strSQL = " SELECT     dbo.RiwayatPekerjaan.NoUrut, dbo.RiwayatPekerjaan.NamaPerusahaan, dbo.RiwayatPekerjaan.JabatanPosisi, " & _
    " dbo.RiwayatPekerjaan.UraianPekerjaan, dbo.RiwayatPekerjaan.TglMulai, dbo.RiwayatPekerjaan.TglAkhir, dbo.RiwayatPekerjaan.GajiPokok," & _
    " dbo.RiwayatPekerjaan.NoSK, dbo.RiwayatPekerjaan.TglSK, dbo.RiwayatPekerjaan.TandaTanganSK, dbo.RiwayatPekerjaan.AlamatPerusahaan," & _
    " dbo.RiwayatPekerjaan.PimpinanPerusahaan, dbo.DataPegawai.NamaLengkap AS NamaUser" & _
    " FROM dbo.RiwayatPekerjaan INNER JOIN" & _
    " dbo.DataPegawai ON dbo.RiwayatPekerjaan.IdUser = dbo.DataPegawai.IdPegawai where RiwayatPekerjaan.IdPegawai = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatPekerjaan.DataSource = rs
    With dgRiwayatPekerjaan
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Alignment = vbCenter
        .Columns("NoUrut").Caption = "No. Urut"
        .Columns("NamaPerusahaan").Width = 2500
        .Columns("NamaPerusahaan").Caption = "Perusahaan"
        .Columns("UraianPekerjaan").Caption = "Uraian Pekerjaan"
        .Columns("JabatanPosisi").Width = 2000
        .Columns("JabatanPosisi").Caption = "Posisi"
        .Columns("TglMulai").Width = 2000
        .Columns("TglMulai").Caption = "Tgl. Mulai"
        .Columns("TglAkhir").Width = 1600
        .Columns("TglAkhir").Caption = "Tgl. Akhir"
        .Columns("GajiPokok").Width = 2000
        .Columns("GajiPokok").NumberFormat = "#,##"
        .Columns("GajiPokok").Caption = "Gaji Pokok"

        .Columns("NoSK").Width = 1600
        .Columns("NoSK").Caption = "No. SK"
        .Columns("TglSK").Caption = "Tgl. SK"
        .Columns("TandaTanganSK").Caption = "TTD SK"
        .Columns("AlamatPerusahaan").Width = 2500
        .Columns("AlamatPerusahaan").Caption = "Alamat Perusahaan"
        .Columns("PimpinanPerusahaan").Caption = "Pimpinan Perusahaan"
        .Columns("NamaUser").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

'Untuk meload riwayat gaji
Public Sub subLoadRiwayatGaji()
    On Error GoTo hell
    strSQL = " SELECT IdPegawai,NoUrut, NoSK, TglSK, TandaTanganSK, KdKomponenGaji,KomponenGaji, Jumlah, Keterangan, NamaUser" & _
    " FROM V_RiwayatGaji where IdPegawai = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatGaji.DataSource = rs
    With dgRiwayatGaji
        .Columns("IdPegawai").Width = 0
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Caption = "No. Urut"
        .Columns("NoSK").Width = 1500
        .Columns("NoSK").Caption = "No. SK"
        .Columns("TglSK").Width = 1200
        .Columns("TglSK").Caption = "Tgl. SK"
        .Columns("TandaTanganSK").Width = 2000
        .Columns("TandaTanganSK").Caption = "TTD SK"
        .Columns("KdKomponenGaji").Width = 0
        .Columns("KomponenGaji").Width = 1600
        .Columns("KomponenGaji").Caption = "Komponen Gaji"
        .Columns("Jumlah").Width = 1200
        .Columns("Jumlah").NumberFormat = "#,###"
        .Columns("Keterangan").Width = 3000
        .Columns("NamaUser").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub subLoadRiwayatStatus()
    On Error GoTo errLoad
    strSQL = " SELECT ID, Nama, Status, [Tgl. Awal], [Tgl. Akhir], [Alasan Keperluan], [Keterangan], NoRiwayat, KdStatus" & _
    " FROM V_RiwayatStatusPegawai_New where ID = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatSta.DataSource = rs
    dgRiwayatSta.Columns("NoRiwayat").Width = 0
    dgRiwayatSta.Columns("KdStatus").Width = 0
    dgRiwayatSta.Columns("Nama").Width = 2000
    dgRiwayatSta.Columns("Alasan Keperluan").Width = 3000
    dgRiwayatSta.Columns("Keterangan").Width = 3000
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Sub subLoadRiwayatPotongan()
    On Error GoTo bawah
    Set rs = Nothing
    strSQL = "SELECT tglPotongan as [Tanggal Potongan], KomponenPotonganGaji as [Komponen Potongan], JumlahPotongan as [Jumlah Potongan], statusEnabled as [Status Aktif], Keterangan FROM V_RiwayatPotonganGaji where idPegawai = '" & txtIdPegawai.Text & "'"
    Call msubRecFO(rs, strSQL)
    Set dgRiwayatPotonganGaji.DataSource = rs
    Exit Sub
bawah:
End Sub

Public Sub subLoadRiwayatTugas()
    On Error GoTo hell
    strSQL = "Select * from V_RiwayatTugas where IdPegawai= '" & mstrIdPegawai & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgTugas.DataSource = rs
    With dgTugas
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Alignment = vbCenter
        .Columns("NoUrut").Caption = "No. Urut"
        .Columns("IdPegawai").Width = 0
        .Columns("NamaTugas").Width = 2150
        .Columns("NamaTugas").Caption = "Tugas"
        .Columns("TglMulai").Width = 1500
        .Columns("TglMulai").Caption = "Tgl. Mulai"
        .Columns("TglAkhir").Width = 1500
        .Columns("TglAkhir").Caption = "Tgl. Akhir"
        .Columns("Alamat").Width = 2500
        .Columns("Alamat").Caption = "Alamat Lokasi"
        .Columns("Keterangan").Width = 3500
        .Columns("NIP").Width = 0
        .Columns("NamaPangkat").Width = 0
        .Columns("NamaGolongan").Width = 0
        .Columns("NamaJabatan").Width = 0
        .Columns("NamaLengkap").Width = 2000
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub subLoadRiwayatSIP()
    On Error GoTo hell
    strSQL = "Select * from RiwayatIzinPraktek where IdPegawai= '" & mstrIdPegawai & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatSIP.DataSource = rs

    Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub subLoadRiwayatMutasiKepegawaian()
'    On Error GoTo hell
'    strSQL = "Select Jabatan, Tempat, Tahun from RiwayatMutasiPegawai where IdPegawai= '" & mstrIdPegawai & "'"
'    Set rs = Nothing
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'    Set dgRiwayatMutasi.DataSource = rs

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDataPegawaiNew.Enabled = True
End Sub

Private Sub sstTP_Click(PreviousTab As Integer)
    Call subLoadRiwayatPendidikanFormal
    Call subLoadRiwayatPendidikanNonFormal
    Call subLoadRiwayatOrganisasi
    Call subLoadRiwayatPerjalananDinas
    Call subLoadRiwayatHukuman
    Call subLoadRiwayatPangkat
    Call subLoadRiwayatExtraPelatihan
    Call subLoadRiwayatPrestasi
    Call subLoadRiwayatPekerjaan
    Call subLoadRiwayatGaji
    Call subLoadRiwayatTempatBertugas
    Call subLoadRiwayatKeluargaPegawai
    Call subLoadRiwayatStatus
    Call subLoadRiwayatTugas
    Call subLoadRiwayatPotongan
    Call subLoadRiwayatSIP
    Call subLoadRiwayatMutasiKepegawaian
End Sub

