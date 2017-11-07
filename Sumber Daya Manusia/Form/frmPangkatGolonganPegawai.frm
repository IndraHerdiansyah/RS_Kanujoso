VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPangkatGolonganPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pangkat & Golongan Pegawai"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   Icon            =   "frmPangkatGolonganPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   8550
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
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
      Left            =   1320
      TabIndex        =   34
      Top             =   7440
      Width           =   1335
   End
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
      Left            =   5640
      TabIndex        =   18
      Top             =   7440
      Width           =   1335
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
      Left            =   4200
      TabIndex        =   17
      Top             =   7455
      Width           =   1335
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
      Left            =   7080
      TabIndex        =   16
      Top             =   7440
      Width           =   1335
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
      Left            =   2760
      TabIndex        =   15
      Top             =   7455
      Width           =   1335
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   6135
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Pangkat"
      TabPicture(0)   =   "frmPangkatGolonganPegawai.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Golongan"
      TabPicture(1)   =   "frmPangkatGolonganPegawai.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   5415
         Left            =   -74760
         TabIndex        =   9
         Top             =   600
         Width           =   8055
         Begin VB.TextBox txtParameter 
            Appearance      =   0  'Flat
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
            Left            =   4560
            MaxLength       =   50
            TabIndex        =   37
            Top             =   5000
            Width           =   3255
         End
         Begin VB.TextBox txtKdExtPangkat 
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
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   26
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox chkStsPangkat 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6600
            TabIndex        =   25
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtNmExtPangkat 
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
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   24
            Top             =   2040
            Width           =   6255
         End
         Begin VB.TextBox txtUrutPangkat 
            Appearance      =   0  'Flat
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
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   2
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtPangkat 
            Appearance      =   0  'Flat
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
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   1
            Top             =   600
            Width           =   6255
         End
         Begin VB.TextBox txtKdPangkat 
            Appearance      =   0  'Flat
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
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   0
            Top             =   240
            Width           =   975
         End
         Begin MSDataGridLib.DataGrid dgPangkat 
            Height          =   2400
            Left            =   240
            TabIndex        =   3
            Top             =   2520
            Width           =   7560
            _ExtentX        =   13335
            _ExtentY        =   4233
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
         Begin MSDataListLib.DataCombo dcGolongan 
            Height          =   315
            Left            =   1560
            TabIndex        =   22
            Top             =   960
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cari Nama Pangkat"
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
            Left            =   2640
            TabIndex        =   38
            Top             =   5040
            Width           =   1515
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
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
            Index           =   6
            Left            =   240
            TabIndex        =   28
            Top             =   1680
            Width           =   1140
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
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
            Index           =   7
            Left            =   240
            TabIndex        =   27
            Top             =   2040
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Golongan"
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
            TabIndex        =   23
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label2 
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
            Left            =   240
            TabIndex        =   21
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label Label12 
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
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label Label13 
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
      End
      Begin VB.Frame Frame2 
         Height          =   5415
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   8055
         Begin VB.TextBox txtParameterGol 
            Appearance      =   0  'Flat
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
            Left            =   4920
            MaxLength       =   20
            TabIndex        =   39
            Top             =   5000
            Width           =   2895
         End
         Begin VB.TextBox txtPPH 
            Alignment       =   1  'Right Justify
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
            Left            =   6840
            MaxLength       =   50
            TabIndex        =   35
            Text            =   "0"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtKdExtGol 
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
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   31
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox chkStatusGol 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6600
            TabIndex        =   30
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtNmExtGol 
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1680
            Width           =   3615
         End
         Begin VB.TextBox txtUrutGol 
            Appearance      =   0  'Flat
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
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   7
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtKdGol 
            Appearance      =   0  'Flat
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
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtGol 
            Appearance      =   0  'Flat
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
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   6
            Top             =   600
            Width           =   5655
         End
         Begin MSDataGridLib.DataGrid dgGol 
            Height          =   2760
            Left            =   240
            TabIndex        =   19
            Top             =   2160
            Width           =   7560
            _ExtentX        =   13335
            _ExtentY        =   4868
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cari Nama Golongan"
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
            Left            =   3000
            TabIndex        =   40
            Top             =   5040
            Width           =   1620
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Persen Pph Jasa"
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
            Index           =   12
            Left            =   5400
            TabIndex        =   36
            Top             =   1680
            Width           =   1320
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
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
            Index           =   4
            Left            =   240
            TabIndex        =   33
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
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
            Index           =   5
            Left            =   240
            TabIndex        =   32
            Top             =   1680
            Width           =   1170
         End
         Begin VB.Label Label1 
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
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nama Golongan"
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
            TabIndex        =   14
            Top             =   600
            Width           =   1275
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
            TabIndex        =   13
            Top             =   240
            Width           =   420
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmPangkatGolonganPegawai.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPangkatGolonganPegawai.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPangkatGolonganPegawai.frx":444B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPangkatGolonganPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub subDCSource()
    strSQL = "SELECT KdGolongan, NamaGolongan FROM GolonganPegawai where StatusEnabled='1' order by NamaGolongan"
    Call msubDcSource(dcGolongan, rs, strSQL)
End Sub

Private Function sp_simpan(f_Status As String) As Boolean
    Select Case sstDataPenunjang.Tab
        Case 0 ' Jenis jabatan
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdPangkat", adVarChar, adParamInput, 2, Trim(txtKdPangkat.Text))
                .Parameters.Append .CreateParameter("NamaPangkat", adVarChar, adParamInput, 50, Trim(txtPangkat.Text))
                .Parameters.Append .CreateParameter("KdGolongan", adVarChar, adParamInput, 2, Trim(dcGolongan.BoundText))
                .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Trim(txtUrutPangkat.Text))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKdExtPangkat.Text)
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNmExtPangkat.Text)
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsPangkat.Value)
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
                MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
            End If
            cmdBatal_Click

        Case 1 ' jabatan
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdGolongan", adVarChar, adParamInput, 2, Trim(txtKdGol.Text))
                .Parameters.Append .CreateParameter("NamaGolongan", adVarChar, adParamInput, 20, Trim(txtGol.Text))
                .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Trim(txtUrutGol.Text))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKdExtGol.Text)
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNmExtGol.Text)
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStatusGol.Value)
                '.Parameters.Append .CreateParameter("PphJasa", adDouble, adParamInput, , txtPPH.Text)
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
            sp_simpanPphGol ("A") '//yayang.agus 2014-08-11
            If (f_Status = "A") Then
                MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
            End If
            cmdBatal_Click

    End Select
End Function

'//yayang.agus 2014-08-11
Private Function sp_simpanPphGol(f_Status As String) As Boolean
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdGolongan", adVarChar, adParamInput, 2, Trim(txtKdGol.Text))
        .Parameters.Append .CreateParameter("Pph", adDouble, adParamInput, 20, Trim(txtPPH.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_GolonganPph"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            MsgBox "Ada kesalahan", vbExclamation, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)

    End With
'    If (f_status = "A") Then
'        MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
'    Else
'        MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
'    End If
'    cmdBatal_Click
End Function
'//


Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadGridSource
    Call subDCSource
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    '    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0
            If dgPangkat.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakPangkat.Show
        Case 1
            If dgGol.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakGolongan.Show
    End Select
hell:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo hell
    Select Case sstDataPenunjang.Tab
        Case 0
            If Periksa("text", txtPangkat, "Silahkan isi nama pangkat ") = False Then Exit Sub
            If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            strSQL = "Select * from DataCurrentPegawai where KdPangkat='" & txtKdPangkat & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa di hapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                Set rs = Nothing
                strSQL = "delete Pangkat where KdPangkat = '" & txtKdPangkat & "'"
                rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
                Set rs = Nothing
            End If
            
            

        Case 1
            If Periksa("text", txtGol, "Silahkan isi nama golongan ") = False Then Exit Sub
            If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            strSQL = "Select * from Pangkat where KdGolongan='" & txtKdGol & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa di hapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                Set rs = Nothing
                strSQL = "delete GolonganPegawai where KdGolongan = '" & txtKdGol & "'"
                rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
                Set rs = Nothing
            End If

    End Select
    Call cmdBatal_Click
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    Select Case sstDataPenunjang.Tab
        Case 0
            If dcGolongan.Text <> "" Then
                If Periksa("datacombo", dcGolongan, "Golongan Tidak Terdaftar") = False Then Exit Sub
            End If
            
            If Periksa("text", txtPangkat, "Silahkan isi nama pangkat ") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub
        Case 1
            If Periksa("text", txtGol, "Silahkan isi nama golongan ") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub
            
    End Select
    Call cmdBatal_Click
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcGolongan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtUrutPangkat.SetFocus

On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcGolongan.Text)) = 0 Then txtUrutPangkat.SetFocus: Exit Sub
        If dcGolongan.MatchedWithList = True Then txtUrutPangkat.SetFocus: Exit Sub
        strSQL = "SELECT KdGolongan, NamaGolongan FROM GolonganPegawai WHERE (NamaGolongan LIKE '%" & dcGolongan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcGolongan.BoundText = rs(0).Value
        dcGolongan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgGol_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgGol
    WheelHook.WheelHook dgGol
End Sub

Private Sub dgGol_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdGol.Text = dgGol.Columns(0).Value
    txtGol.Text = dgGol.Columns(1).Value
    If IsNull(dgGol.Columns(2)) Then txtUrutGol.Text = "" Else txtUrutGol.Text = dgGol.Columns(2)
    txtKdExtGol.Text = dgGol.Columns(3).Value
    txtNmExtGol.Text = dgGol.Columns(4).Value
    chkStatusGol.Value = dgGol.Columns(5).Value
    If dgGol.Columns(6) = "" Then txtPPH.Text = 0 Else txtPPH.Text = dgGol.Columns(6).Value
End Sub

Private Sub dgPangkat_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPangkat
    WheelHook.WheelHook dgPangkat
End Sub

Private Sub dgPangkat_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdPangkat.Text = dgPangkat.Columns(0).Value
    txtPangkat.Text = dgPangkat.Columns(1).Value
    If IsNull(dgPangkat.Columns(3)) Then dcGolongan.BoundText = "" Else dcGolongan.BoundText = dgPangkat.Columns(3)
    If IsNull(dgPangkat.Columns(2)) Then txtUrutPangkat.Text = "" Else txtUrutPangkat.Text = dgPangkat.Columns(2)
    txtKdExtPangkat.Text = dgPangkat.Columns(5).Value
    txtNmExtPangkat.Text = dgPangkat.Columns(6).Value
    chkStsPangkat.Value = dgPangkat.Columns(7).Value
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    sstDataPenunjang.Tab = 0
    Call subDCSource
    Call cmdBatal_Click
End Sub

Sub subKosong()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0 'Jenis jabatan
            txtKdPangkat.Text = ""
            txtPangkat.Text = ""
            txtUrutPangkat.Text = ""
            dcGolongan.BoundText = ""
            txtKdPangkat.SetFocus
            txtKdExtPangkat = ""
            txtNmExtPangkat = ""
            chkStsPangkat.Value = 1
        Case 1 'jabatan
            txtKdGol.Text = ""
            txtGol.Text = ""
            txtUrutGol.Text = ""
            txtKdGol.SetFocus
            txtKdExtGol = ""
            txtNmExtGol = ""
            chkStatusGol.Value = 1
    End Select
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call cmdBatal_Click
End Sub

Private Sub txtKdExtGol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtGol.SetFocus
End Sub

Private Sub txtKdExtPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtPangkat.SetFocus
End Sub

Private Sub txtKdGol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtGol.SetFocus
End Sub

Private Sub txtKdPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPangkat.SetFocus
End Sub

Private Sub txtNmExtGol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPPH.SetFocus
End Sub

Private Sub txtNmExtPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcGolongan.SetFocus
End Sub

Sub subLoadGridSource()
    On Error GoTo hell
    Select Case sstDataPenunjang.Tab
        Case 0
            Set rs = Nothing
            strSQL = "SELECT dbo.Pangkat.KdPangkat AS Kode, dbo.Pangkat.NamaPangkat AS [Nama Pangkat], dbo.Pangkat.NoUrut AS [No. Urut], dbo.Pangkat.KdGolongan, " & _
            "dbo.GolonganPegawai.NamaGolongan AS Golongan,dbo.Pangkat.KodeExternal,dbo.Pangkat.NamaExternal,dbo.Pangkat.StatusEnabled " & _
            "FROM dbo.Pangkat LEFT OUTER JOIN " & _
            "dbo.GolonganPegawai ON dbo.Pangkat.KdGolongan = dbo.GolonganPegawai.KdGolongan where dbo.Pangkat.NamaPangkat LIKE '%" & txtParameter.Text & "%'  ORDER BY dbo.Pangkat.NoUrut "
            Call msubRecFO(rs, strSQL)
            Set dgPangkat.DataSource = rs
            dgPangkat.Columns(1).Width = 4500
            dgPangkat.Columns(2).Width = 1000
            dgPangkat.Columns(3).Width = 0
            dgPangkat.Columns(4).Width = 1000
            dgPangkat.Columns(7).Width = 1250
        Case 1
            '//yayang.agus 2014-08-11
            Set rs = Nothing
            strSQL = "select KdGolongan AS Kode, NamaGolongan AS [Nama Golongan], NoUrut AS [No. Urut],KodeExternal,NamaExternal,StatusEnabled,Pph from V_GolonganPegawai where NamaGolongan LIKE '%" & txtParameterGol.Text & "%' "
            Call msubRecFO(rs, strSQL)
            Set dgGol.DataSource = rs
            dgGol.Columns(0).Width = 1500
            dgGol.Columns(0).Alignment = vbCenter
            dgGol.Columns(1).Width = 3800
            dgGol.Columns(2).Width = 1000
            dgGol.Columns(5).Width = 1250
            dgGol.Columns(6).Width = 1000
            '//
'            Set rs = Nothing
'            strSQL = "select KdGolongan AS Kode, NamaGolongan AS [Nama Golongan], NoUrut AS [No. Urut],KodeExternal,NamaExternal,StatusEnabled from GolonganPegawai where NamaGolongan LIKE '%" & txtParameterGol.Text & "%' "
'            Call msubRecFO(rs, strSQL)
'            Set dgGol.DataSource = rs
'            dgGol.Columns(0).Width = 1500
'            dgGol.Columns(0).Alignment = vbCenter
'            dgGol.Columns(1).Width = 3800
'            dgGol.Columns(2).Width = 1000
'            dgGol.Columns(5).Width = 1250
    End Select
    Exit Sub
hell:
'    Call msubPesanError
End Sub

Private Sub txtGol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtUrutGol.SetFocus
End Sub

Private Sub txtParameter_Change()
    Call subLoadGridSource
    strCetak = " where dbo.Pangkat.NamaPangkat LIKE '%" & txtParameter.Text & "%'"
End Sub

Private Sub txtParameterGol_Change()
    Call subLoadGridSource
    strCetak = " where NamaGolongan LIKE '%" & txtParameter.Text & "%'"
End Sub

Private Sub txtPPH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtPPH_LostFocus()
    If txtPPH.Text = "" Then txtPPH.Text = 0
End Sub

Private Sub txtUrutGol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtGol.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtUrutPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtPangkat.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub
