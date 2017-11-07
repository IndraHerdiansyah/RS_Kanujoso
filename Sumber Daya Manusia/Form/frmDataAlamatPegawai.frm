VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmDataAlamatPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Alamat Pegawai"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmDataAlamatPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10695
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
      Left            =   7800
      TabIndex        =   20
      Top             =   6480
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
      Left            =   6360
      TabIndex        =   19
      Top             =   6480
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
      Left            =   9240
      TabIndex        =   21
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
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
      Left            =   4920
      TabIndex        =   18
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox txtnourut 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   600
      MaxLength       =   20
      TabIndex        =   42
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Pegawai"
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
      Left            =   120
      TabIndex        =   31
      Top             =   960
      Width           =   10455
      Begin VB.TextBox txtIdPegawai 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNamaLengkap 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtJenisPegawai 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtJabatan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   7320
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   4680
         MaxLength       =   1
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. ID Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Lengkap"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   35
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "JK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4680
         TabIndex        =   34
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   33
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7320
         TabIndex        =   32
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Alamat Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   10455
      Begin VB.ComboBox cbStatusAktif 
         Height          =   315
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtstatus 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   9600
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtMail 
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
         Height          =   330
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtFax 
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
         Height          =   330
         Left            =   5160
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtHp 
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
         Height          =   330
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtAlamat 
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
         Height          =   330
         Left            =   120
         MaxLength       =   100
         TabIndex        =   5
         Top             =   600
         Width           =   9375
      End
      Begin VB.TextBox txtKodePos 
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
         Height          =   330
         Left            =   120
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtTelp 
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
         Height          =   330
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1800
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo dcKecamatan 
         Height          =   315
         Left            =   5280
         TabIndex        =   9
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSDataListLib.DataCombo dcKota 
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSDataListLib.DataCombo dcPropinsi 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1200
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
      Begin MSDataListLib.DataCombo dcKelurahan 
         Height          =   315
         Left            =   7800
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSMask.MaskEdBox meRTRW 
         Height          =   330
         Left            =   9600
         TabIndex        =   6
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         HideSelection   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##/##"
         Mask            =   "##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   9480
         TabIndex        =   41
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7080
         TabIndex        =   40
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Faximile"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   39
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "HP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   38
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Kelurahan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7800
         TabIndex        =   30
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Lengkap"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "RT/RW"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9600
         TabIndex        =   28
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Kode Pos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Propinsi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Kota / Kabupaten"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   25
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Kecamatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   24
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Telepon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   23
         Top             =   1560
         Width           =   570
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   37
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
   Begin MSDataGridLib.DataGrid dgalamat 
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataAlamatPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8880
      Picture         =   "frmDataAlamatPegawai.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataAlamatPegawai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDataAlamatPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vbMsgboxRslt As VbMsgBoxResult

Sub SetComboPropinsi()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from Propinsi where StatusEnabled='1' ", dbConn, , adLockOptimistic
    Set dcPropinsi.RowSource = rs
    dcPropinsi.ListField = rs.Fields(1).Name
    dcPropinsi.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Exit Sub
hell:
End Sub

Private Sub cbStatusAktif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub cmdBatal_Click()
    Call kosong
    txtAlamat.SetFocus
    cmdSimpan.Enabled = True
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus pegawai dengan NIP '" _
    & txtIdPegawai.Text & "'" & vbNewLine _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbYes Then
        If txtNoUrut.Text = "" Then Exit Sub
        strSQL = "DELETE FROM DataAlamatPegawai WHERE IdPegawai='" & txtIdPegawai.Text & "' AND NoUrut='" & txtNoUrut.Text & "'"
        dbConn.Execute strSQL
        Call loadgrid
        Call kosong
        MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    End If
    Exit Sub
errHapus:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    If dcKecamatan.Text <> "" Then
        If Periksa("datacombo", dcKecamatan, "Kecamatan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcKelurahan.Text <> "" Then
        If Periksa("datacombo", dcKelurahan, "Kelurahan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcKota.Text <> "" Then
        If Periksa("datacombo", dcKota, "Kota Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcPropinsi.Text <> "" Then
        If Periksa("datacombo", dcPropinsi, "Provinsi Tidak Terdaftar") = False Then Exit Sub
    End If

    If Periksa("text", txtAlamat, "Alamat harus diisi!") = False Then Exit Sub
    'If Periksa("text", txtstatus, "Status Aktif Alamat Harus Diisi!") = False Then Exit Sub
    If cbStatusAktif.Text = "" Then
        MsgBox "Status Aktif Alamat Harus Diisi!", vbExclamation, "Validasi"
        Exit Sub
    End If
    
    If sp_simpan = False Then Exit Sub

    Call kosong
    Call loadgrid
    Call cmdBatal_Click
End Sub

Private Function sp_simpan() As Boolean
    On Error GoTo hell
    sp_simpan = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIdPegawai.Text)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, Null)
        End If
        .Parameters.Append .CreateParameter("AlamatLengkap", adVarChar, adParamInput, 100, txtAlamat.Text)
        .Parameters.Append .CreateParameter("Kelurahan", adVarChar, adParamInput, 50, IIf(dcKelurahan.Text = "", Null, dcKelurahan.Text))
        .Parameters.Append .CreateParameter("Kecamatan", adVarChar, adParamInput, 50, IIf(dcKecamatan.Text = "", Null, dcKecamatan.Text))
        .Parameters.Append .CreateParameter("KotaKabupaten", adVarChar, adParamInput, 50, IIf(dcKota.Text = "", Null, dcKota.Text))
        .Parameters.Append .CreateParameter("Propinsi", adVarChar, adParamInput, 50, IIf(dcPropinsi.Text = "", Null, dcPropinsi.Text))
        .Parameters.Append .CreateParameter("RTRW", adVarChar, adParamInput, 7, IIf(meRTRW.Text = "", Null, meRTRW.Text))
        .Parameters.Append .CreateParameter("KodePos", adChar, adParamInput, 5, IIf(txtKodePos.Text = "", Null, txtKodePos.Text))
        .Parameters.Append .CreateParameter("Telepon", adVarChar, adParamInput, 20, IIf(txtTelp.Text = "", Null, txtTelp.Text))
        .Parameters.Append .CreateParameter("Handphone", adVarChar, adParamInput, 30, IIf(txtHp.Text = "", Null, txtHp.Text))
        .Parameters.Append .CreateParameter("Faximile", adVarChar, adParamInput, 20, IIf(txtFax.Text = "", Null, txtFax.Text))
        .Parameters.Append .CreateParameter("Email", adVarChar, adParamInput, 50, IIf(txtMail.Text = "", Null, txtMail.Text))
        .Parameters.Append .CreateParameter("StatusAktif", adChar, adParamInput, 1, cbStatusAktif.Text)
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 3, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_DataAlamatPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_simpan = False
        Else
            If Not IsNull(.Parameters("OutputNoUrut").Value) Then txtNoUrut = .Parameters("OutputNoUrut").Value
            mstrIdPegawai = txtIdPegawai.Text
            MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Sub cmdTutup_Click()
    frmDataPegawaiNew.Enabled = True
    Unload Me
End Sub

Private Sub dcKecamatan_GotFocus()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from Kecamatan where KdPropinsi='" & dcPropinsi.BoundText & "' and KdKotaKabupaten='" & dcKota.BoundText & "'", dbConn, , adLockOptimistic
    Set dcKecamatan.RowSource = rs
    dcKecamatan.ListField = rs.Fields(3).Name
    dcKecamatan.BoundColumn = rs.Fields(2).Name
    Set rs = Nothing
    Exit Sub
hell:
End Sub

Private Sub dcKecamatan_KeyPress(KeyAscii As Integer)
On Error GoTo Errload
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKecamatan.Text)) = 0 Then dcKelurahan.SetFocus: Exit Sub
        If dcKecamatan.MatchedWithList = True Then dcKelurahan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "Select KdPropinsi, KdKotaKabupaten, KdKecamatan, NamaKecamatan from Kecamatan where KdPropinsi='" & dcPropinsi.BoundText & "' AND KdKotaKabupaten='" & dcKota.BoundText & "' AND NamaKecamatan LIKE '%" & dcKecamatan.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcKecamatan.BoundText = dbRst(2).Value
        dcKecamatan.Text = dbRst(3).Value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcKelurahan_Click(Area As Integer)
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "select * from Kelurahan where KdPropinsi='" & dcPropinsi.BoundText & "' and KdKotaKabupaten='" & dcKota.BoundText & "' and KdKecamatan='" & dcKecamatan.BoundText & "' and KdKelurahan='" & dcKelurahan.BoundText & "'", dbConn, adOpenStatic, adLockReadOnly
    If rs.RecordCount <> 0 Then
        txtKodePos.Text = rs.Fields("KodePos").Value
    End If
    Set rs = Nothing
    Exit Sub
hell:
End Sub

Private Sub dcKelurahan_GotFocus()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from Kelurahan where KdPropinsi='" & dcPropinsi.BoundText & "' and KdKotaKabupaten='" & dcKota.BoundText & "' and KdKecamatan='" & dcKecamatan.BoundText & "'", dbConn, , adLockOptimistic
    Set dcKelurahan.RowSource = rs
    dcKelurahan.ListField = rs.Fields(5).Name
    dcKelurahan.BoundColumn = rs.Fields(3).Name
    Set rs = Nothing
    Exit Sub
hell:
End Sub

Private Sub dcKelurahan_KeyPress(KeyAscii As Integer)
On Error GoTo Errload
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKelurahan.Text)) = 0 Then dcKelurahan.SetFocus: Exit Sub
        If dcKelurahan.MatchedWithList = True Then dcKelurahan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "Select * from Kelurahan where KdPropinsi='" & dcPropinsi.BoundText & "' and KdKotaKabupaten='" & dcKota.BoundText & "' and KdKecamatan='" & dcKecamatan.BoundText & "' AND NamaKelurahan LIKE '%" & dcKelurahan.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcKelurahan.BoundText = dbRst(3).Value
        dcKelurahan.Text = dbRst(5).Value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcKota_GotFocus()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from KotaKabupaten where KdPropinsi='" & dcPropinsi.BoundText & "'", dbConn, , adLockOptimistic
    Set dcKota.RowSource = rs
    dcKota.ListField = rs.Fields(2).Name
    dcKota.BoundColumn = rs.Fields(1).Name
    Set rs = Nothing
    Exit Sub
hell:
End Sub

Private Sub dcKota_KeyPress(KeyAscii As Integer)
On Error GoTo Errload
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKota.Text)) = 0 Then dcKecamatan.SetFocus: Exit Sub
        If dcKota.MatchedWithList = True Then dcKecamatan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "Select KdPropinsi, KdKotaKabupaten, NamaKotaKabupaten from KotaKabupaten where KdPropinsi='" & dcPropinsi.BoundText & "' AND NamaKotaKabupaten LIKE '%" & dcKota.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcKota.BoundText = dbRst(1).Value
        dcKota.Text = dbRst(2).Value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcPropinsi_KeyPress(KeyAscii As Integer)
On Error GoTo Errload
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcPropinsi.Text)) = 0 Then dcKota.SetFocus: Exit Sub
        If dcPropinsi.MatchedWithList = True Then dcKota.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "Select KdPropinsi, NamaPropinsi from Propinsi where NamaPropinsi LIKE '%" & dcPropinsi.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcPropinsi.BoundText = dbRst(0).Value
        dcPropinsi.Text = dbRst(1).Value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dgalamat_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgalamat.ApproxCount = 0 Then Exit Sub
    With dgalamat
        txtNoUrut.Text = .Columns("NoUrut").Value
        txtIdPegawai.Text = .Columns("ID").Value
        txtNamaLengkap.Text = .Columns("Nama").Value
        txtJK.Text = .Columns("JK").Value
        txtJenisPegawai.Text = .Columns("Jenis Pegawai").Value
        If IsNull(.Columns("Jabatan")) Then txtJabatan.Text = "" Else txtJabatan.Text = .Columns("Jabatan").Value
        txtAlamat.Text = .Columns("Alamat Lengkap").Value
        If IsNull(.Columns("Kelurahan")) Then dcKelurahan.Text = "" Else dcKelurahan.Text = .Columns("Kelurahan").Value
        If IsNull(.Columns("Kecamatan")) Then dcKecamatan.Text = "" Else dcKecamatan.Text = .Columns("Kecamatan").Value
        If IsNull(.Columns("Kota/Kabupaten")) Then dcKota.Text = "" Else dcKota.Text = .Columns("Kota/Kabupaten").Value
        If IsNull(.Columns("Propinsi")) Then dcPropinsi.Text = "" Else dcPropinsi.Text = .Columns("Propinsi").Value
        If IsNull(.Columns("RT/RW")) Then meRTRW.Text = "__/__" Else meRTRW.Text = .Columns("RT/RW").Value
        If IsNull(.Columns("Kode Pos")) Then txtKodePos.Text = "" Else txtKodePos.Text = .Columns("Kode Pos").Value
        If IsNull(.Columns("Telepon")) Then txtTelp.Text = "" Else txtTelp.Text = .Columns("Telepon").Value
        If IsNull(.Columns("Hand Phone")) Then txtHp.Text = "" Else txtHp.Text = .Columns("Hand Phone").Value
        If IsNull(.Columns("Faximile")) Then txtFax.Text = "" Else txtFax.Text = .Columns("Faximile").Value
        If IsNull(.Columns("E-mail")) Then txtMail.Text = "" Else txtMail.Text = .Columns("E-mail").Value
        txtstatus.Text = .Columns("Status Aktif").Value
        cbStatusAktif.Text = .Columns("Status Aktif").Value
    End With
End Sub

Private Sub Form_Load()
    With frmDataPegawaiNew
        txtIdPegawai.Text = .txtIdPegawai.Text
        txtNamaLengkap.Text = .txtNama.Text
        txtJK.Text = .cbJK.Text
        txtJenisPegawai.Text = .dcJnsPeg.Text
        txtJabatan.Text = .dcJabatan.Text
    End With
    
    cbStatusAktif.AddItem "Y"
    cbStatusAktif.AddItem "T"
    
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call SetComboPropinsi
    Call loadgrid
End Sub

Sub kosong()
    On Error Resume Next
    txtNoUrut.Text = ""
    txtAlamat.Text = ""
    dcKelurahan.Text = ""
    dcKecamatan.Text = ""
    dcKota.Text = ""
    dcPropinsi.Text = ""
    meRTRW.Text = "__/__"
    txtKodePos.Text = ""
    txtTelp.Text = ""
    txtHp.Text = ""
    txtFax.Text = ""
    txtMail.Text = ""
    txtstatus.Text = ""
    cbStatusAktif.Text = ""
    txtAlamat.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDataPegawaiNew.Enabled = True
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then meRTRW.SetFocus
End Sub

Private Sub meRTRW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dcPropinsi.SetFocus
End Sub

Private Sub dcPropinsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKota.SetFocus
End Sub

Private Sub dcKota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKecamatan.SetFocus
End Sub

Private Sub dcKecamatan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKelurahan.SetFocus
End Sub

Private Sub dcKelurahan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKodePos.SetFocus
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtMail.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtHp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFax.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtKodePos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtTelp.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtMail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbStatusAktif.SetFocus
End Sub

Private Sub txtstatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtTelp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtHp.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Sub loadgrid()
    On Error GoTo hell
    Set rs = Nothing
    strSQL = "select * from v_AlamatPegawai WHERE ID='" & mstrIdPegawai & "' ORDER BY NoUrut"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgalamat.DataSource = rs
    dgalamat.Columns("NoUrut").Width = 0
    Exit Sub
hell:
    Call msubPesanError
End Sub
