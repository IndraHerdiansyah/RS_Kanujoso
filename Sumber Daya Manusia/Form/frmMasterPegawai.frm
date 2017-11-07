VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Data Pegawai "
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   Icon            =   "frmMasterPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Kelompok Pegawai"
      TabPicture(0)   =   "frmMasterPegawai.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Jenis Pegawai"
      TabPicture(1)   =   "frmMasterPegawai.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Golongan Pegawai"
      TabPicture(2)   =   "frmMasterPegawai.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pangkat Pegawai"
      TabPicture(3)   =   "frmMasterPegawai.frx":0D1E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Jenis Jabatan"
      TabPicture(4)   =   "frmMasterPegawai.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Jabatan"
      TabPicture(5)   =   "frmMasterPegawai.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame3"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   44
         Top             =   840
         Width           =   7575
         Begin VB.TextBox txtGolPegawai 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   50
            TabIndex        =   47
            Top             =   480
            Width           =   5415
         End
         Begin VB.TextBox txtKodeGol 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            MaxLength       =   2
            TabIndex        =   46
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtNoUrutGol 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6480
            MaxLength       =   2
            TabIndex        =   45
            Top             =   480
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid dgGolonganPegawai 
            Height          =   3135
            Left            =   240
            TabIndex        =   48
            Top             =   960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Golongan Pegawai"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   960
            TabIndex        =   51
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   50
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "No. Urut"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6480
            TabIndex        =   49
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   36
         Top             =   840
         Width           =   7575
         Begin VB.TextBox txtKdJenisPegawai 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            MaxLength       =   3
            TabIndex        =   38
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtJenisPegawai 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3360
            MaxLength       =   50
            TabIndex        =   37
            Top             =   480
            Width           =   3975
         End
         Begin MSDataListLib.DataCombo dcKelompokPegawai 
            Height          =   330
            Left            =   1080
            TabIndex        =   39
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgJenisPegawai 
            Height          =   3135
            Left            =   240
            TabIndex        =   40
            Top             =   960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   43
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Pegawai"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3360
            TabIndex        =   42
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kelompok Pegawai"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1080
            TabIndex        =   41
            Top             =   240
            Width           =   1530
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   30
         Top             =   840
         Width           =   7575
         Begin VB.TextBox txtKelompokPegawai 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   50
            TabIndex        =   32
            Top             =   600
            Width           =   6375
         End
         Begin VB.TextBox txtKdKelompokPegawai 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            MaxLength       =   2
            TabIndex        =   31
            Top             =   600
            Width           =   615
         End
         Begin MSDataGridLib.DataGrid dgKelompokPegawai 
            Height          =   3015
            Left            =   240
            TabIndex        =   33
            Top             =   1080
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5318
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kelompok Pegawai"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   960
            TabIndex        =   35
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   420
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   20
         Top             =   840
         Width           =   7575
         Begin VB.TextBox txtNamaJabatan 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   23
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtKdJabatan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            MaxLength       =   5
            TabIndex        =   22
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtNoUrutJabatan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   21
            Top             =   480
            Width           =   735
         End
         Begin MSDataListLib.DataCombo dcJenisJabatan 
            Height          =   330
            Left            =   4680
            TabIndex        =   24
            Top             =   480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgJabatan 
            Height          =   3135
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nama Jabatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1200
            TabIndex        =   28
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Jabatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4680
            TabIndex        =   27
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "No. Urut"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6600
            TabIndex        =   26
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   14
         Top             =   840
         Width           =   7575
         Begin VB.TextBox txtKdJenisJabatan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            MaxLength       =   2
            TabIndex        =   15
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtJenisJabatan 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   16
            Top             =   480
            Width           =   6015
         End
         Begin MSDataGridLib.DataGrid dgJenisJabatan 
            Height          =   3135
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Kode Jenis"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Jabatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1320
            TabIndex        =   18
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4335
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   7575
         Begin VB.TextBox txtNmPangkat 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   50
            TabIndex        =   9
            Top             =   480
            Width           =   5295
         End
         Begin VB.TextBox txtKodePangkat 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            MaxLength       =   2
            TabIndex        =   8
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtNoUrutPangkat 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            MaxLength       =   2
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
         Begin MSDataGridLib.DataGrid dgPangkat 
            Height          =   3135
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nama Pangkat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   960
            TabIndex        =   12
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "No. Urut"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6360
            TabIndex        =   10
            Top             =   240
            Width           =   705
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
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
      Left            =   6550
      Picture         =   "frmMasterPegawai.frx":0D72
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterPegawai.frx":1AFA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterPegawai.frx":44BB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMasterPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub subDcSource()
    On Error GoTo errLoad
    'Jenis pegawai
    strSQL = "SELECT * FROM KelompokPegawai order by KelompokPegawai"
    Call msubDcSource(dcKelompokPegawai, rs, strSQL)

    strSQL = "SELECT * FROM JenisJabatan order by JenisJabatan"
    Call msubDcSource(dcJenisJabatan, rs, strSQL)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_simpan(f_Status As String) As Boolean
    On Error GoTo errLoad
    Select Case sstDataPenunjang.Tab
        Case 0 'kelompok pegawai
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdKelompokPegawai", adChar, adParamInput, 2, Trim(txtKdKelompokPegawai))
                .Parameters.Append .CreateParameter("KelompokPegawai", adVarChar, adParamInput, 50, Trim(txtKelompokPegawai))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_KelompokPegawai"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            If (f_Status = "A") Then
                MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
            End If
            cmdBatal_Click

        Case 1 'Jenis Pegawai
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdJenisPegawai", adChar, adParamInput, 3, Trim(txtKdJenisPegawai.Text))
                .Parameters.Append .CreateParameter("KdKelompokPegawai", adChar, adParamInput, 2, Trim(dcKelompokPegawai.BoundText))
                .Parameters.Append .CreateParameter("JenisPegawai", adVarChar, adParamInput, 50, Trim(txtJenisPegawai.Text))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_JenisPegawai"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            If (f_Status = "A") Then
                MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
            End If
            cmdBatal_Click

        Case 2 'Golongan Pegawai
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdGolongan", adVarChar, adParamInput, 2, Trim(txtKodeGol.Text))
                .Parameters.Append .CreateParameter("NamaGolongan", adVarChar, adParamInput, 20, Trim(txtGolPegawai.Text))
                .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Trim(txtNoUrutGol.Text))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_GolonganPegawai"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            If (f_Status = "A") Then
                MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
            End If
            cmdBatal_Click

        Case 3 'pangkat pegawai
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdPangkat", adVarChar, adParamInput, 2, Trim(txtKodePangkat.Text))
                .Parameters.Append .CreateParameter("NamaPangkat", adVarChar, adParamInput, 50, Trim(txtNmPangkat.Text))
                .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Trim(txtNoUrutPangkat.Text))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_Pangkat"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)
            End With
            If (f_Status = "A") Then
                MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
            End If
            cmdBatal_Click

        Case 4 ' Jenis jabatan
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdJenisJabatan", adVarChar, adParamInput, 2, Trim(txtKdJenisJabatan))
                .Parameters.Append .CreateParameter("JenisJabatan", adVarChar, adParamInput, 30, Trim(txtJenisJabatan))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_JenisJabatan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            If (f_Status = "A") Then
                MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
            End If
            cmdBatal_Click

        Case 5 ' jabatan
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdJabatan", adVarChar, adParamInput, 5, Trim(txtKdJabatan))
                .Parameters.Append .CreateParameter("NamaJabatan", adVarChar, adParamInput, 50, Trim(txtNamaJabatan))
                .Parameters.Append .CreateParameter("KdJenisJabatan", adVarChar, adParamInput, 2, Trim(dcJenisJabatan.BoundText))
                .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Trim(txtNoUrutJabatan.Text))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_Jabatan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            If (f_Status = "A") Then
                MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
            End If
            cmdBatal_Click

    End Select
    Exit Function
errLoad:
    Call msubPesanError
End Function

Private Sub cmdBatal_Click()
    Call subLoadGridSource
    Select Case sstDataPenunjang.Tab
        Case 0 'kelompok pegawai
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 1 'Jenis Pegawai
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 2 'Golongan Pegawai
            txtKodeGol.Enabled = True
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 3 'pangkat pegawai
            txtKodePangkat.Enabled = True
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 4 ' jabatan
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 5 ' Jenis jabatan
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
    End Select
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo hell
    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case sstDataPenunjang.Tab
        Case 0 'kelompok pegawai
            If Periksa("text", txtKelompokPegawai, "Pilih Data yang akan dihapus") = False Then Exit Sub
            If sp_simpan("D") = False Then Exit Sub

        Case 1 'jenis pegawai
            If Periksa("datacombo", dcKelompokPegawai, "Pilih Data yang akan dihapus") = False Then Exit Sub
            If sp_simpan("D") = False Then Exit Sub

        Case 2  ' golongan pegawai
            If Periksa("text", txtKodeGol, "Pilih Data yang akan dihapus") = False Then Exit Sub
            If sp_simpan("D") = False Then Exit Sub

        Case 3 'pangkat pegawai
            If Periksa("text", txtKodePangkat, "Pilih Data yang akan dihapus") = False Then Exit Sub
            If sp_simpan("D") = False Then Exit Sub
            '
        Case 4 'Jenis jabatan
            If Periksa("text", txtJenisJabatan, "Pilih Data yang akan dihapus") = False Then Exit Sub
            If sp_simpan("D") = False Then Exit Sub

        Case 5 ' jabatan
            If Periksa("datacombo", dcJenisJabatan, "Pilih Data yang akan dihapus") = False Then Exit Sub
            If Periksa("text", txtNamaJabatan, "Isi Nama jabatan") = False Then Exit Sub
            If sp_simpan("D") = False Then Exit Sub
    End Select
    MsgBox "Data berhasil dihapus", vbInformation, "Informasi"

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan
    Select Case sstDataPenunjang.Tab
        Case 0 'kelompok pegawai
            If Periksa("text", txtKelompokPegawai, "Isi Kelompok pegawai") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub

        Case 1 'jenis pegawai
            If Periksa("datacombo", dcKelompokPegawai, "Isi Kelompok pegawai!") = False Then Exit Sub
            If Periksa("text", txtJenisPegawai, "Isi Jenis pegawai!") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub

        Case 2  ' golongan pegawai
            If Periksa("text", txtKodeGol, "Kode Golongan Harus diisi!!") = False Then Exit Sub
            If Periksa("text", txtGolPegawai, "Isi Golongan pegawai!") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub

        Case 3 'pangkat pegawai
            If Periksa("text", txtKodePangkat, "Kode Pangkat Harus diisi!!") = False Then Exit Sub
            If Periksa("text", txtNmPangkat, "Isi Pangkat pegawai!") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub
            '
        Case 4 'Jenis jabatan
            If Periksa("text", txtJenisJabatan, "Isi jenis jabatan!") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub

        Case 5 ' jabatan
            If Periksa("datacombo", dcJenisJabatan, "Isi Jenis Jabatan!") = False Then Exit Sub
            If Periksa("text", txtNamaJabatan, "Isi Nama jabatan") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub
    End Select
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
    Exit Sub
errSimpan:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoUrutJabatan.SetFocus
End Sub

Private Sub dcKelompokPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisPegawai.SetFocus
End Sub

Private Sub dgGolonganPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKodeGol.Text = dgGolonganPegawai.Columns(0).Value
    txtGolPegawai.Text = dgGolonganPegawai.Columns(1).Value
    txtNoUrutGol.Text = dgGolonganPegawai.Columns(2).Value
    txtKodeGol.Enabled = False
End Sub

Private Sub dgJabatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdJabatan.Text = dgJabatan.Columns(0).Value
    dcJenisJabatan.Text = IIf(dgJabatan.Columns(2).Value = Null, "", dgJabatan.Columns(2).Value)
    txtNamaJabatan.Text = dgJabatan.Columns(1).Value
    txtNoUrutJabatan.Text = dgJabatan.Columns(3).Value
    '    txtKodeJenis.Enabled = False
End Sub

Private Sub dgJenisJabatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdJenisJabatan.Text = dgJenisJabatan.Columns(0).Value
    txtJenisJabatan.Text = dgJenisJabatan.Columns(1).Value
End Sub

Private Sub dgJenisPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdJenisPegawai.Text = dgJenisPegawai.Columns("Kode").Text
    dcKelompokPegawai.Text = dgJenisPegawai.Columns("Kelompok Pegawai").Text
    txtJenisPegawai.Text = dgJenisPegawai.Columns("Jenis Pegawai").Text
End Sub

Private Sub dgKelompokPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdKelompokPegawai.Text = dgKelompokPegawai.Columns(0).Value
    txtKelompokPegawai.Text = dgKelompokPegawai.Columns(1).Value
End Sub

Private Sub dgPangkat_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKodePangkat.Text = dgPangkat.Columns(0).Value
    txtNmPangkat.Text = dgPangkat.Columns(1).Value
    txtNoUrutPangkat.Text = dgPangkat.Columns(2).Value
    txtKodePangkat.Enabled = False
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subDcSource
    sstDataPenunjang.Tab = 0
    Call subLoadGridSource
End Sub

Sub subKosong()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0 'Kelompok Pegawai
            txtKdKelompokPegawai.Text = ""
            txtKelompokPegawai.Text = ""
            txtKelompokPegawai.SetFocus

        Case 1 'Jenis pegawai
            txtKdJenisPegawai.Text = ""
            '            txtkdkelompok.Text = ""
            txtJenisPegawai.Text = ""
            dcKelompokPegawai.Text = ""
            dcKelompokPegawai.SetFocus

        Case 2 'Golongan Pegawai
            txtKodeGol.Text = ""
            txtGolPegawai.Text = ""
            txtNoUrutGol.Text = ""
            txtKodeGol.SetFocus

        Case 3 'Pangkat Pegawai
            txtKodePangkat.Text = ""
            txtNmPangkat.Text = ""
            txtNoUrutPangkat.Text = ""
            txtKodePangkat.SetFocus
        Case 4 'Jenis jabatan
            txtKdJenisJabatan.Text = ""
            txtJenisJabatan.Text = ""
            txtJenisJabatan.SetFocus

        Case 5 'jabatan
            txtKdJabatan.Text = ""
            dcJenisJabatan.Text = ""
            txtNamaJabatan.Text = ""
            txtNoUrutJabatan.Text = ""
            dcJenisJabatan.SetFocus
    End Select
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call subDcSource
    Call subLoadGridSource
    Call cmdBatal_Click
End Sub

Private Sub txtGolPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoUrutGol.SetFocus
End Sub

Private Sub txtJenisJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtJenisPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKdJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaJabatan.SetFocus
End Sub

Private Sub txtKelompokPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKodeGol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtGolPegawai.SetFocus
End Sub

Private Sub txtKodePangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmPangkat.SetFocus
End Sub

Sub subLoadGridSource()
    On Error GoTo errLoad
    Select Case sstDataPenunjang.Tab
        Case 0 ' kelompok pegawai
            Set rs = Nothing
            strSQL = "select * from kelompokpegawai"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgKelompokPegawai.DataSource = rs
            dgKelompokPegawai.Columns(0).DataField = rs(0).Name
            dgKelompokPegawai.Columns(1).DataField = rs(1).Name
            dgKelompokPegawai.Columns(0).Width = 1250
            dgKelompokPegawai.Columns(0).Caption = "Kode"
            dgKelompokPegawai.Columns(1).Width = 5500
            dgKelompokPegawai.Columns(1).Caption = "Nama Kelompok"
            Set rs = Nothing

        Case 1  'jenis pegawai
            Set rs = Nothing
            strSQL = "select * from v_y_jenispegawai "
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With dgJenisPegawai
                Set .DataSource = rs

                .Columns.Item(0).Visible = False
                .Columns(1).Width = 1250
                .Columns(2).Width = 2000
                .Columns(3).Width = 3200
            End With
            Set rs = Nothing

        Case 2  'golongan pegawai
            Set rs = Nothing
            strSQL = "select * from golonganpegawai"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgGolonganPegawai.DataSource = rs
            dgGolonganPegawai.Columns(0).DataField = rs(0).Name
            dgGolonganPegawai.Columns(1).DataField = rs(1).Name
            dgGolonganPegawai.Columns(0).Width = 1250
            dgGolonganPegawai.Columns(0).Caption = "Kode"
            dgGolonganPegawai.Columns(1).Width = 4200
            dgGolonganPegawai.Columns(1).Caption = "Nama Golongan"
            dgGolonganPegawai.Columns(2).Width = 1000
            dgGolonganPegawai.Columns(2).Caption = "No. Urut"
            Set rs = Nothing

        Case 3 'pangkat pegawai
            Set rs = Nothing
            strSQL = "select * from pangkat"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgPangkat.DataSource = rs
            dgPangkat.Columns(0).DataField = rs(0).Name
            dgPangkat.Columns(1).DataField = rs(1).Name
            dgPangkat.Columns(0).Width = 1250
            dgPangkat.Columns(0).Caption = "Kode"
            dgPangkat.Columns(1).Width = 4200
            dgPangkat.Columns(1).Caption = "Nama Pangkat"
            dgPangkat.Columns(2).Caption = "No. Urut"
            dgPangkat.Columns(2).Width = 1000
            Set rs = Nothing

        Case 4 ' Jenis jabatan
            Set rs = Nothing
            strSQL = "select * from JenisJabatan"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJenisJabatan.DataSource = rs
            dgJenisJabatan.Columns(0).DataField = rs(0).Name
            dgJenisJabatan.Columns(1).DataField = rs(1).Name
            dgJenisJabatan.Columns(0).Width = 1250
            dgJenisJabatan.Columns(0).Caption = "Kode"
            dgJenisJabatan.Columns(1).Width = 5500
            dgJenisJabatan.Columns(1).Caption = "Jenis Jabatan"
            Set rs = Nothing

        Case 5  'Jabatan
            Set rs = Nothing
            strSQL = "SELECT dbo.Jabatan.KdJabatan, dbo.Jabatan.NamaJabatan, dbo.JenisJabatan.JenisJabatan, dbo.Jabatan.NoUrut" & _
            " FROM dbo.Jabatan LEFT OUTER JOIN" & _
            " dbo.JenisJabatan ON dbo.Jabatan.KdJenisJabatan = dbo.JenisJabatan.KdJenisJabatan"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJabatan.DataSource = rs
            dgJabatan.Columns(0).DataField = rs(0).Name
            dgJabatan.Columns(1).DataField = rs(1).Name
            dgJabatan.Columns(2).DataField = rs(2).Name
            dgJabatan.Columns(0).Width = 1250
            dgJabatan.Columns(0).Caption = "Kode"
            dgJabatan.Columns(1).Width = 3900
            dgJabatan.Columns(1).Caption = "Nama Jabatan"
            dgJabatan.Columns(2).Width = 1800
            dgJabatan.Columns(2).Caption = "Jenis Jabatan"
            dgJabatan.Columns(3).Caption = "No. Urut"
            dgJabatan.Columns(3).Width = 1000
            Set rs = Nothing
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub txtNamaJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisJabatan.SetFocus
End Sub

Private Sub txtNmPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoUrutPangkat.SetFocus
End Sub

Private Sub txtNoUrutGol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNoUrutJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNoUrutPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

