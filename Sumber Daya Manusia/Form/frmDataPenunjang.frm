VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Begin VB.Form frmDataPenunjang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Penunjang"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmDataPenunjang.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7635
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   7575
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Baru"
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
         Left            =   195
         TabIndex        =   7
         Top             =   225
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
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
         Left            =   6000
         TabIndex        =   6
         Top             =   240
         Width           =   1335
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
         Left            =   4575
         TabIndex        =   5
         Top             =   255
         Width           =   1335
      End
      Begin VB.CommandButton cmdUbah 
         Caption         =   "&Ubah"
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
         Left            =   3135
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
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
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   4575
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   5
      TabHeight       =   520
      TabCaption(0)   =   "Kelompok Pegawai"
      TabPicture(0)   =   "frmDataPenunjang.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Jenis Pegawai"
      TabPicture(1)   =   "frmDataPenunjang.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Golongan Pegawai"
      TabPicture(2)   =   "frmDataPenunjang.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pangkat Pegawai"
      TabPicture(3)   =   "frmDataPenunjang.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Komponen Gaji"
      TabPicture(4)   =   "frmDataPenunjang.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   35
         Top             =   840
         Width           =   7095
         Begin VB.TextBox txtKodeKelompok 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            MaxLength       =   2
            TabIndex        =   0
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtKelompokGaji 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   1
            Top             =   600
            Width           =   3975
         End
         Begin MSDataGridLib.DataGrid dgKomponenGaji 
            Height          =   2175
            Left            =   240
            TabIndex        =   36
            Top             =   1080
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
               Caption         =   "KODE GAJI"
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
               Caption         =   "KELOMPOK GAJI"
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kelompok"
            Height          =   210
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Kelompok Gaji"
            Height          =   195
            Left            =   1680
            TabIndex        =   37
            Top             =   360
            Width           =   1020
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3495
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   7215
         Begin VB.TextBox txtKelPegawai 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   31
            Top             =   600
            Width           =   4335
         End
         Begin VB.TextBox txtKode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            MaxLength       =   2
            TabIndex        =   30
            Top             =   615
            Width           =   1215
         End
         Begin MSDataGridLib.DataGrid dgKelompokPegawai 
            Height          =   2055
            Left            =   240
            TabIndex        =   32
            Top             =   1080
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
               Caption         =   "KODE PEGAWAI"
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
               Caption         =   "KELOMPOK PEGAWAI"
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kelompok Pegawai"
            Height          =   210
            Left            =   1680
            TabIndex        =   34
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kelompok"
            Height          =   210
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   21
         Top             =   840
         Width           =   7095
         Begin VB.TextBox txtJenisPegawai 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4440
            MaxLength       =   50
            TabIndex        =   23
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtKodeJenis 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            MaxLength       =   3
            TabIndex        =   22
            Top             =   600
            Width           =   855
         End
         Begin MSDataListLib.DataCombo dcKelompok 
            Height          =   330
            Left            =   1200
            TabIndex        =   24
            Top             =   600
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataGridLib.DataGrid dgJenisPegawai 
            Height          =   2175
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   "KD. KELOMPOK"
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
               Caption         =   "KELOMPOK"
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
            BeginProperty Column02 
               DataField       =   ""
               Caption         =   "KD. JENIS"
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
            BeginProperty Column03 
               DataField       =   ""
               Caption         =   "JENIS"
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kelompok Pegawai"
            Height          =   210
            Left            =   1200
            TabIndex        =   28
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Pegawai"
            Height          =   210
            Left            =   4440
            TabIndex        =   27
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Kode Jenis"
            Height          =   210
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   870
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   15
         Top             =   840
         Width           =   7215
         Begin VB.TextBox txtGolPegawai 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   17
            Top             =   600
            Width           =   4455
         End
         Begin VB.TextBox txtKodeGol 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            MaxLength       =   2
            TabIndex        =   16
            Top             =   600
            Width           =   1215
         End
         Begin MSDataGridLib.DataGrid dgGolonganPegawai 
            Height          =   2175
            Left            =   240
            TabIndex        =   18
            Top             =   1080
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
               Caption         =   "KODE GOLONGAN"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   "GOLONGAN PEGAWAI"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Golongan Pegawai"
            Height          =   210
            Left            =   1680
            TabIndex        =   20
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Kode Golongan"
            Height          =   210
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   9
         Top             =   840
         Width           =   7215
         Begin VB.TextBox txtNmPangkat 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   11
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox txtKodePangkat 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            MaxLength       =   2
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin MSDataGridLib.DataGrid dgPangkat 
            Height          =   2175
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   16
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
               Caption         =   "KODE PANGKAT"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   "NAMA PANGKAT"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nama Pangkat"
            Height          =   210
            Left            =   1680
            TabIndex        =   14
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Kode Pangkat"
            Height          =   210
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1140
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   39
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataPenunjang.frx":0D56
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   5760
      Picture         =   "frmDataPenunjang.frx":3717
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataPenunjang.frx":4C05
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDataPenunjang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 'kelompok pegawai
            txtkode.Enabled = True
            clear
            cmdHapus.Enabled = False
            cmdubah.Enabled = False
            cmdsimpan.Enabled = True
        Case 1 'Jenis Pegawai
            txtKodeJenis.Enabled = True
            clear
            cmdHapus.Enabled = False
            cmdubah.Enabled = False
            cmdsimpan.Enabled = True
        Case 2 'Golongan Pegawai
            txtKodeGol.Enabled = True
            clear
            cmdHapus.Enabled = False
            cmdubah.Enabled = False
            cmdsimpan.Enabled = True
        Case 3 'pangkat pegawai
            txtKodePangkat.Enabled = True
            clear
            cmdHapus.Enabled = False
            cmdubah.Enabled = False
            cmdsimpan.Enabled = True
        Case 4 'Komponen Gaji
            txtKodeKelompok.Enabled = True
            clear
            cmdHapus.Enabled = False
            cmdubah.Enabled = False
            cmdsimpan.Enabled = True
    End Select
End Sub

Private Sub cmdHapus_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 'kelompok pegawai
            Set rs = Nothing
            strSQL = "select * from jenispegawai where kdkelompokpegawai = '" & txtkode & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "Data tidak dapat Di hapus"
                dgKelompokPegawai.SetFocus
                Exit Sub
            End If
            
            Set rs = Nothing
            strSQL = "delete kelompokpegawai  where kdkelompokpegawai = '" & txtkode & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 1 'jenis pegawai
            Set rs = Nothing
            strSQL = "select * from datapegawai where kdjenispegawai = '" & txtKodeJenis & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox " data tidak dapat dihapus"
                dgJenisPegawai.SetFocus
                Exit Sub
            End If
            Set rs = Nothing
            strSQL = "delete jenispegawai where kdjenispegawai= '" & txtKodeJenis & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 2 'golongan pegawai
            Set rs = Nothing
            strSQL = "select * from datacurrentpegawai where kdgolongan = '" & txtKodeGol & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "data tidak dapat Di Hapus"
                dgGolonganPegawai.SetFocus
                Exit Sub
            End If
            Set rs = Nothing
            strSQL = "delete golonganpegawai where kdgolongan = '" & txtKodeGol & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 3 'pangkat pegawai
            Set rs = Nothing
            strSQL = "select * from datapegawai where kdpangkat = '" & txtKodePangkat & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
               MsgBox " data tidak dapat dihapus"
               dgPangkat.SetFocus
               Exit Sub
            End If
            Set rs = Nothing
            strSQL = "delete pangkat where kdpangkat = '" & txtKodePangkat & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 4 'Komponen Gaji
            Set rs = Nothing
            strSQL = "delete komponengaji where kdkomponengaji = '" & txtKodeKelompok & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
    End Select
    tampilData
End Sub

Private Sub cmdSimpan_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 'kelompok pegawai
            Set rs = Nothing
            strSQL = "select * from kelompokpegawai where kdkelompokpegawai = '" & txtkode & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "Kode Pegawai Sudah Ada"
                txtkode.SetFocus
                Exit Sub
            End If
            Set rs = Nothing
            strSQL = "insert into kelompokpegawai(kdkelompokpegawai,kelompokpegawai) " & _
                     "values ('" & txtkode & "', '" & txtKelPegawai & "')"
            
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 1 'jenis pegawai
            Set rs = Nothing
            strSQL = "select * from jenispegawai where kdjenispegawai  = '" & txtKodeJenis & "' "
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
               MsgBox "Kode Jenis Sudah Ada"
               txtKodeJenis.SetFocus
               Exit Sub
            End If
            Set rs = Nothing
            strSQL = "insert into jenispegawai (kdjenispegawai,kdkelompokpegawai,jenispegawai)" & _
                     "values ('" & txtKodeJenis & "','" & dcKelompok.BoundText & "','" & txtJenisPegawai & "')"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 2  ' golongan pegawai
            Set rs = Nothing
            strSQL = "select * from golonganpegawai where kdgolongan = '" & txtKodeGol & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
               MsgBox "Kode Golongan Sudah Ada"
               txtKodeGol.SetFocus
               Exit Sub
            End If
            Set rs = Nothing
            strSQL = "insert into golonganpegawai (kdgolongan, namagolongan)" & _
                     "values ('" & txtKodeGol & "', '" & txtGolPegawai & "')"
            
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 3 'pangkat pegawai
            Set rs = Nothing
            strSQL = "select * from pangkat where kdpangkat = '" & txtKodePangkat & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
               MsgBox "Kode Pangkat sudah ada"
               txtKodePangkat.SetFocus
               Exit Sub
            End If
            Set rs = Nothing
            strSQL = "insert into pangkat (kdpangkat,namapangkat)" & _
                     "values('" & txtKodePangkat & "','" & txtNmPangkat & "')"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 4 'komponen gaji
            Set rs = Nothing
            strSQL = "select * from komponengaji where kdkomponengaji = '" & txtKodeKelompok & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
               MsgBox "Kode Komponen Gaji sudah ada"
               txtKodeKelompok.SetFocus
               Exit Sub
            End If
            Set rs = Nothing
            strSQL = "insert into komponengaji (kdkomponengaji,komponengaji)" & _
                     "values('" & txtKodeKelompok & "','" & txtKelompokGaji & "')"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
    End Select
    tampilData
End Sub

Private Sub cmdsimpan_GotFocus()
    Select Case sstDataPenunjang.Tab
        Case 0   'kelompok pegawai
            If txtKelPegawai.Text = "" Then
                MsgBox "Kelompok pegawai harap diisi"
                txtKelPegawai.SetFocus
                Exit Sub
            End If
        Case 1 'jenis pegawai
            If dcKelompok.Text = "" Then
                MsgBox "Kelompok Pegawai Harap diisi"
                dcKelompok.SetFocus
                Exit Sub
            End If
        Case 2 'golongan pegawai
            If txtGolPegawai.Text = "" Then
               MsgBox "Golongan Pegawai Harap Diisi"
               txtGolPegawai.SetFocus
               Exit Sub
            End If
        Case 3 'pangkat pegawai
            If txtNmPangkat.Text = "" Then
               MsgBox "Nama Pangkat harap Diisi"
               txtNmPangkat.SetFocus
               Exit Sub
            End If
        Case 4 'komponen gaji
            If txtKelompokGaji.Text = "" Then
               MsgBox "Nama Kelompok Gaji harap Diisi"
               txtKelompokGaji.SetFocus
               Exit Sub
            End If
    End Select
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdubah_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 ' kelompok pegawai
            Set rs = Nothing
            strSQL = "update kelompokpegawai set kelompokpegawai = '" & txtKelPegawai & "' where kdkelompokpegawai = '" & txtkode & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 1 'jenispegawai
            Set rs = Nothing
            strSQL = "update jenispegawai set jenispegawai = '" & txtJenisPegawai & "' where kdjenispegawai = '" & txtKodeJenis & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 2 'golongan pegawai
            Set rs = Nothing
            strSQL = "update golonganpegawai set NamaGolongan = '" & txtGolPegawai & "' where kdgolongan = '" & txtKodeGol & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 3 ' pangkat pegawai
            Set rs = Nothing
            strSQL = "update pangkat set namapangkat = '" & txtNmPangkat & "' where kdpangkat = '" & txtKodePangkat & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 4 ' komponen gaji
            Set rs = Nothing
            strSQL = "update komponengaji set komponengaji = '" & txtKelompokGaji & "' where kdkomponengaji = '" & txtKodeKelompok & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
    End Select
    tampilData
End Sub

Sub SetComboKelompokPegawai()
    Set rs = Nothing
    rs.Open "Select * from kelompokpegawai", dbConn, , adLockOptimistic
    Set dcKelompok.RowSource = rs
    dcKelompok.ListField = rs.Fields(1).Name
    dcKelompok.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Private Sub dcKelompok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisPegawai.SetFocus
End Sub

Private Sub dgGolonganPegawai_Click()
    cmdsimpan.Enabled = False
    cmdubah.Enabled = True
    cmdHapus.Enabled = True
    txtKodeGol.Text = dgGolonganPegawai.Columns(0).Value
    txtGolPegawai.Text = dgGolonganPegawai.Columns(1).Value
    txtKodeGol.Enabled = False
End Sub

Private Sub dgGolonganPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtKodeGol.Text = dgGolonganPegawai.Columns(0).Value
    txtGolPegawai.Text = dgGolonganPegawai.Columns(1).Value
    txtKodeGol.Enabled = False
End Sub

Private Sub dgJenisPegawai_Click()
    cmdsimpan.Enabled = False
    cmdubah.Enabled = True
    cmdHapus.Enabled = True
    txtKodeJenis.Text = dgJenisPegawai.Columns(2).Value
    txtJenisPegawai.Text = dgJenisPegawai.Columns(3).Value
    dcKelompok.Text = dgJenisPegawai.Columns(1).Value
    txtKodeJenis.Enabled = False
End Sub

Private Sub dgJenisPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtKodeJenis.Text = dgJenisPegawai.Columns(2).Value
    txtJenisPegawai.Text = dgJenisPegawai.Columns(3).Value
    dcKelompok.Text = dgJenisPegawai.Columns(1).Value
    txtKodeJenis.Enabled = False
End Sub

Private Sub dgKelompokPegawai_Click()
    cmdsimpan.Enabled = False
    cmdubah.Enabled = True
    cmdHapus.Enabled = True
    txtkode.Text = dgKelompokPegawai.Columns(0).Value
    txtKelPegawai.Text = dgKelompokPegawai.Columns(1).Value
    txtkode.Enabled = False
End Sub

Private Sub dgKelompokPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtkode.Text = dgKelompokPegawai.Columns(0).Value
    txtKelPegawai.Text = dgKelompokPegawai.Columns(1).Value
    txtkode.Enabled = False
End Sub

Private Sub dgPangkat_Click()
    cmdsimpan.Enabled = False
    cmdubah.Enabled = True
    cmdHapus.Enabled = True
    txtKodePangkat.Text = dgPangkat.Columns(0).Value
    txtNmPangkat.Text = dgPangkat.Columns(1).Value
    txtKodePangkat.Enabled = False
End Sub

Private Sub dgPangkat_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtKodePangkat.Text = dgPangkat.Columns(0).Value
    txtNmPangkat.Text = dgPangkat.Columns(1).Value
    txtKodePangkat.Enabled = False
End Sub

Private Sub dgKomponenGaji_Click()
    cmdsimpan.Enabled = False
    cmdubah.Enabled = True
    cmdHapus.Enabled = True
    txtKodeKelompok.Text = dgKomponenGaji.Columns(0).Value
    txtKelompokGaji.Text = dgKomponenGaji.Columns(1).Value
    txtKodeKelompok.Enabled = False
End Sub

Private Sub dgKomponenGaji_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtKodeKelompok.Text = dgKomponenGaji.Columns(0).Value
    txtKelompokGaji.Text = dgKomponenGaji.Columns(1).Value
    txtKodeKelompok.Enabled = False
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    cmdHapus.Enabled = False
    cmdubah.Enabled = False

    'case 0 'kelompok pegawai
    Set rs = Nothing
    strSQL = "select * from kelompokpegawai"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgKelompokPegawai.DataSource = rs
        dgKelompokPegawai.Columns(0).DataField = rs(0).Name
        dgKelompokPegawai.Columns(1).DataField = rs(1).Name
        dgKelompokPegawai.Columns(0).Width = 2000
        dgKelompokPegawai.Columns(1).Width = 2500
        dgKelompokPegawai.ReBind
    Set rs = Nothing

    'case 1 'jenis pegawai
    Set rs = Nothing
    strSQL = "select * from V_Y_jenispegawai"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgJenisPegawai.DataSource = rs
        dgJenisPegawai.Columns(0).DataField = rs(0).Name
        dgJenisPegawai.Columns(1).DataField = rs(1).Name
        dgJenisPegawai.Columns(2).DataField = rs(2).Name
        dgJenisPegawai.Columns(3).DataField = rs(3).Name
        dgJenisPegawai.ReBind
    Set rs = Nothing
    Call SetComboKelompokPegawai

    'Case 2  'golongan pegawai
    Set rs = Nothing
    strSQL = "select * from golonganpegawai"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgGolonganPegawai.DataSource = rs
        dgGolonganPegawai.Columns(0).DataField = rs(0).Name
        dgGolonganPegawai.Columns(1).DataField = rs(1).Name
        dgGolonganPegawai.Columns(1).Width = 2500
        dgGolonganPegawai.ReBind
    Set rs = Nothing

    'case 3 'pangkat pegawai
    Set rs = Nothing
    strSQL = "select * from pangkat"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgPangkat.DataSource = rs
        dgPangkat.Columns(0).DataField = rs(0).Name
        dgPangkat.Columns(1).DataField = rs(1).Name
        dgPangkat.Columns(1).Width = 2500
        dgPangkat.ReBind
    Set rs = Nothing

    'case 4 'komponen gaji
    Set rs = Nothing
    strSQL = "select * from komponengaji"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgKomponenGaji.DataSource = rs
        dgKomponenGaji.Columns(0).DataField = rs(0).Name
        dgKomponenGaji.Columns(1).DataField = rs(1).Name
        dgKomponenGaji.Columns(1).Width = 2500
        dgKomponenGaji.ReBind
    Set rs = Nothing
    
    txtKodeJenis.Enabled = True
End Sub

Sub clear()
    Select Case sstDataPenunjang.Tab
        Case 0 'Kelompok Pegawai
            txtkode.Text = ""
            txtKelPegawai.Text = ""
            txtkode.SetFocus

        Case 1 'Jenis pegawai
            txtKodeJenis.Text = ""
            txtJenisPegawai.Text = ""
            dcKelompok.Text = ""
            txtKodeJenis.SetFocus

        Case 2 'Golongan Pegawai
            txtKodeGol.Text = ""
            txtGolPegawai.Text = ""
            txtKodeGol.SetFocus

        Case 3 'Pangkat Pegawai
            txtKodePangkat.Text = ""
            txtNmPangkat.Text = ""
            txtKodePangkat.SetFocus
        
        Case 4 'Komponen Gaji
            txtKodeKelompok.Text = ""
            txtKelompokGaji.Text = ""
            txtKodeKelompok.SetFocus
    End Select
End Sub

Private Sub txtGolPegawai_GotFocus()
    If txtKodeGol.Text = "" Then
        MsgBox "Nama Golongan Harus Terisi"
        txtGolPegawai.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtJenisPegawai_GotFocus()
    If txtKodeJenis.Text = "" Then
        MsgBox "kode Jenis harap diisi"
        txtkode.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtKelPegawai_GotFocus()
    If txtkode.Text = "" Then
        MsgBox "Kode Harap Diisi"
        txtkode.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKelPegawai.SetFocus
End Sub

Private Sub txtKodeGol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtGolPegawai.SetFocus
End Sub

Private Sub txtKodeJenis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKelompok.SetFocus
End Sub

Private Sub txtKodePangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmPangkat.SetFocus
End Sub

Private Sub txtNmPangkat_GotFocus()
    If txtKodePangkat.Text = "" Then
        MsgBox "Kode Pangkat Harap Terisi"
        txtKodePangkat.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtKodeKelompok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKelompokGaji.SetFocus
End Sub

Private Sub txtKelompokGaji_GotFocus()
    If txtKodeKelompok.Text = "" Then
        MsgBox "Kode Komponen Gaji Harap Terisi"
        txtKodeKelompok.SetFocus
        Exit Sub
    End If
End Sub

Sub tampilData()
    Select Case sstDataPenunjang.Tab
        Case 0 ' kelompok pegawai
            Set rs = Nothing
            strSQL = "select * from kelompokpegawai"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgKelompokPegawai.DataSource = rs
                dgKelompokPegawai.Columns(0).DataField = rs(0).Name
                dgKelompokPegawai.Columns(1).DataField = rs(1).Name
                dgKelompokPegawai.ReBind
            Set rs = Nothing

        Case 1  'jenis pegawai
            Set rs = Nothing
            strSQL = "select * from v_y_jenispegawai "
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJenisPegawai.DataSource = rs
                dgJenisPegawai.Columns(0).DataField = rs(0).Name
                dgJenisPegawai.Columns(1).DataField = rs(1).Name
                dgJenisPegawai.Columns(2).DataField = rs(2).Name
                dgJenisPegawai.Columns(3).DataField = rs(3).Name
                dgJenisPegawai.ReBind
            Set rs = Nothing

        Case 2  'golongan pegawai
            Set rs = Nothing
            strSQL = "select * from golonganpegawai"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgGolonganPegawai.DataSource = rs
                dgGolonganPegawai.Columns(0).DataField = rs(0).Name
                dgGolonganPegawai.Columns(1).DataField = rs(1).Name
                dgGolonganPegawai.ReBind
            Set rs = Nothing

        Case 3 'pangkat pegawai
            Set rs = Nothing
            strSQL = "select * from pangkat"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgPangkat.DataSource = rs
                dgPangkat.Columns(0).DataField = rs(0).Name
                dgPangkat.Columns(1).DataField = rs(1).Name
                dgPangkat.ReBind
            Set rs = Nothing
            
        Case 4 'komponen gaji
            Set rs = Nothing
            strSQL = "select * from komponengaji"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgKomponenGaji.DataSource = rs
                dgKomponenGaji.Columns(0).DataField = rs(0).Name
                dgKomponenGaji.Columns(1).DataField = rs(1).Name
                dgKomponenGaji.ReBind
            Set rs = Nothing
    End Select
End Sub


