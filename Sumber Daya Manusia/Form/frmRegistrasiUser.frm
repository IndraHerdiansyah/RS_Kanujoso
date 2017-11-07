VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegistrasiUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Finger Print Pegawai"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12525
   Icon            =   "frmRegistrasiUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   12525
   Begin MSDataGridLib.DataGrid dgNamaPegawai 
      Height          =   1935
      Left            =   5880
      TabIndex        =   35
      Top             =   4920
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   16
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
   Begin VB.TextBox txtIsi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3600
      TabIndex        =   34
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picUncheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4680
      Picture         =   "frmRegistrasiUser.frx":0CCA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   33
      Top             =   8280
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4440
      Picture         =   "frmRegistrasiUser.frx":0F14
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   32
      Top             =   8280
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   9600
      TabIndex        =   31
      ToolTipText     =   "Simpan ke database"
      Top             =   8280
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid fgRegUser 
      Height          =   3975
      Left            =   120
      TabIndex        =   30
      Top             =   4200
      Width           =   12275
      _ExtentX        =   21643
      _ExtentY        =   7011
      _Version        =   393216
      GridLinesFixed  =   1
      Appearance      =   0
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
   Begin VB.Frame Frame3 
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
      Height          =   3015
      Left            =   8280
      TabIndex        =   9
      Top             =   1080
      Width           =   4095
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtJabatan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtTempatTugas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtNamaPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtIDPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   10
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   2520
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan"
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   210
         Index           =   8
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus Data"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      ToolTipText     =   "Hapus data finger print dari alat Finger Print"
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpanData 
      Caption         =   "Simpan Data"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      ToolTipText     =   "Simpan data finger print dari alat Finger Print"
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdAmbilData 
      Caption         =   "Ambil Data"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Ambil data finger print dari alat Finger Print"
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   11040
      TabIndex        =   18
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   4080
      TabIndex        =   22
      Top             =   1080
      Width           =   4095
      Begin VB.Frame Frame4 
         Caption         =   "Data yang akan diambil"
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   3615
         Begin VB.OptionButton optDataReg 
            Caption         =   "Finger Print"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkFP 
            Appearance      =   0  'Flat
            Caption         =   "Finger Print 3"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   8
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox chkFP 
            Appearance      =   0  'Flat
            Caption         =   "Finger Print 2"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   7
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CheckBox chkFP 
            Appearance      =   0  'Flat
            Caption         =   "Finger Print 1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optDataReg 
            Caption         =   "Password"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox txtNoRegFP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmbHakAkses 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmRegistrasiUser.frx":115E
         Left            =   1560
         List            =   "frmRegistrasiUser.frx":116E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Hak Akses"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "No. Reg FP"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   3855
      Begin MSComctlLib.TreeView trvFP 
         Height          =   2415
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4260
         _Version        =   393217
         Indentation     =   176
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblIP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<lokasi>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   3615
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   19
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
      Left            =   10560
      Picture         =   "frmRegistrasiUser.frx":11B2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRegistrasiUser.frx":1F3A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmRegistrasiUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
