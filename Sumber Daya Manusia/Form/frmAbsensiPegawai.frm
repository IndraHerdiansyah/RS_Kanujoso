VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbsensiPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Absensi Pegawai"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbsensiPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   13995
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   6630
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   22066
            TextSave        =   "24/06/2013"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "15:35"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9720
      TabIndex        =   39
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdpulang 
      Caption         =   "pulang"
      Height          =   375
      Left            =   5880
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer TimerHapusPIN 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5640
      Top             =   1080
   End
   Begin MSDataGridLib.DataGrid dgAbsensi 
      Height          =   2655
      Left            =   0
      TabIndex        =   36
      Top             =   3960
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   4683
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
   Begin VB.TextBox txtkode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtistrahatakhir 
      Height          =   375
      Left            =   5640
      TabIndex        =   34
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtistrahatawal 
      Height          =   375
      Left            =   3840
      TabIndex        =   33
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtakhir 
      Height          =   375
      Left            =   2040
      TabIndex        =   32
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdakhir 
      Caption         =   "akhir"
      Height          =   375
      Left            =   5160
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdawal 
      Caption         =   "awal"
      Height          =   375
      Left            =   4440
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6480
      TabIndex        =   29
      Text            =   "Text2"
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer timerPIN 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6840
      Top             =   1080
   End
   Begin VB.TextBox txtPIN 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      MaxLength       =   5
      TabIndex        =   28
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtkodestatus 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   8520
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Status Kondisi Finger Scan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   25
      Top             =   1080
      Width           =   3615
      Begin VB.ListBox List1 
         Height          =   2370
         ItemData        =   "frmAbsensiPegawai.frx":0CCA
         Left            =   240
         List            =   "frmAbsensiPegawai.frx":0CCC
         TabIndex        =   26
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdmasuk 
      Caption         =   "masuk"
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtnoriwayat 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Absensi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   6375
      Begin VB.Frame Frame5 
         Caption         =   "Cek"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1935
         Left            =   3120
         TabIndex        =   40
         Top             =   480
         Width           =   3015
         Begin VB.CommandButton cmdDisconnect 
            Caption         =   "&Disconnect"
            Height          =   375
            Left            =   360
            TabIndex        =   43
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton cmdconnect 
            Caption         =   "&Connect"
            Height          =   375
            Left            =   360
            TabIndex        =   42
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton cmdsetting 
            Caption         =   "&Setting"
            Height          =   375
            Left            =   360
            MaskColor       =   &H80000000&
            TabIndex        =   41
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Timer tmr_CekError 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   4080
         Top             =   0
      End
      Begin VB.TextBox txttgl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   20
         Top             =   600
         Width           =   2655
      End
      Begin VB.Timer kirim_konfirmasi 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   3600
         Top             =   0
      End
      Begin VB.Timer minta_absensi 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2400
         Top             =   0
      End
      Begin VB.TextBox txtstatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtFingerPrint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Absen"
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finger Scan Client"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   210
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   645
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   19200
      SThreshold      =   1
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   12120
         TabIndex        =   22
         Top             =   4200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   2775
      Left            =   10200
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
      Begin VB.TextBox txtjabatan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txttempattugas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtjk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtIDPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
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
         TabIndex        =   0
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   2520
         TabIndex        =   6
         Top             =   240
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
         TabIndex        =   5
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   210
         Index           =   8
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   975
      Left            =   3000
      TabIndex        =   21
      Top             =   7440
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1720
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgPin 
      Height          =   2175
      Left            =   7680
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12120
      Picture         =   "frmAbsensiPegawai.frx":0CCE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   0
      Picture         =   "frmAbsensiPegawai.frx":1A56
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   1800
      Picture         =   "frmAbsensiPegawai.frx":4417
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmAbsensiPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msg As VbMsgBoxResult

Private Sub cmdakhir_Click()
    On Error Resume Next
    Set rs = Nothing
    strSQL = "SELECT Max(NoRiwayat) as NoRiwayat, TglMasuk, TglPulang" _
    & " FROM AbsensiPegawai" _
    & " WHERE IdPegawai='" & txtIDPegawai.Text & "'" _
    & " GROUP BY TglMasuk, TglPulang"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    If txtStatus.Text = "ISTIRAHAT AKHIR" Then
        rs.MoveLast
        If rs.EOF = False Then
            If Not IsNull(Format(rs.Fields("TglMasuk").Value, "yyyy/MM/dd") = Format(txttgl.Text, "yyyy/MM/dd")) Then
                If Format(rs.Fields("TglPulang").Value, "yyyy/MM/dd") = Format(txttgl.Text, "yyyy/MM/dd") Then
                    GoTo hell
                Else
                    If UpdateAbsensiAkhir("A") = False Then Exit Sub
                    Call txtidpegawai_Change
                End If
            End If
        Else
            GoTo hell
        End If
    End If
    Exit Sub
    Set rs = Nothing
hell:
End Sub

Private Sub cmdawal_Click()
    On Error Resume Next
    Set rs = Nothing
    strSQL = "SELECT Max(NoRiwayat) as NoRiwayat, TglMasuk, TglPulang" _
    & " FROM AbsensiPegawai" _
    & " WHERE IdPegawai='" & txtIDPegawai.Text & "'" _
    & " GROUP BY TglMasuk, TglPulang"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    If txtStatus.Text = "ISTIRAHAT AWAL" Then
        rs.MoveLast
        If rs.EOF = False Then
            If Not IsNull(Format(rs.Fields("TglMasuk").Value, "yyyy/MM/dd")) Then
                If Format(rs.Fields("TglPulang").Value, "yyyy/MM/dd") = Format(txttgl, "yyyy/MM/dd") Then
                    GoTo hell
                Else
                    If UpdateAbsensiAwal("A") = False Then Exit Sub
                    Call txtidpegawai_Change
                End If
            End If
        Else
            GoTo hell
        End If
    End If
    Exit Sub
    Set rs = Nothing
hell:
End Sub

Private Sub cmdConnect_Click()
    'perintah untuk melakukan koneksi serial
    If add = "" Then
        MsgBox "Setting Parameter Koneksi Kurang", vbOKOnly, "Pesan koneksi"
        Exit Sub
    Else
        cmdconnect.Enabled = False
        cmdDisconnect.Enabled = True
        cmdsetting.Enabled = False
        Call Get_Connect
        If MSComm1.PortOpen = True Then
            Call reset1
            Call reset2
            frmKFRS.cmdOk.Enabled = False 'diganti
            frmKFRS.cmdBatal.Enabled = False
            kirim_konfirmasi.Enabled = True
        Else
            cmdconnect.Enabled = True
            cmdsetting.Enabled = True
            cmdDisconnect.Enabled = False
        End If
    End If
End Sub

Private Sub cmdDisconnect_Click()
    'perintah untuk memutus koneksi serial

    i = MsgBox("Putus Koneksi dengan FRS-400!", vbOKCancel, "Putus koneksi")
    If i = vbOK Then
        minta_absensi.Enabled = False
        tmr_CekError.Enabled = False
        Call Get_Disconnect
        cmdconnect.Enabled = True
        cmdsetting.Enabled = True
        cmdDisconnect.Enabled = False
        Call reset1
        Call reset2
        List1.clear
        List1.List(0) = "Tidak Ada Koneksi"
        Call kosong
    End If
End Sub

Private Sub cmdsetting_Click()
    frmSetting.Show vbModal, MDIUtama
End Sub

Private Sub Command1_Click()

    List1.List(1) = "       " & "FRS - 1 (Tidak Beroperasi)"
    List1.List(2) = "       " & "FRS - 2 (Tidak Beroperasi)"
    List1.List(3) = "       " & "FRS - 3 (Tidak Beroperasi)"

End Sub

Private Sub Form_Load()
    add2 = &H0
    Call reset1
    Call reset2
    cmdconnect.Enabled = True
    cmdsetting.Enabled = True
    cmdDisconnect.Enabled = False
    centerForm Me, MDIUtama
    Call PlayFlashMovie(Me)
    Call loadgrid
    Call txtidpegawai_Change
    List1.List(0) = "Tidak Ada Koneksi"
End Sub

Private Sub cmdmasuk_Click()
    On Error Resume Next
    If txtPIN.Text = "" Then Exit Sub
    Set rs = Nothing
    strSQL = "SELECT TglMasuk, TglPulang FROM AbsensiPegawai" _
    & " WHERE IdPegawai='" & txtIDPegawai.Text & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    If txtStatus.Text = "MASUK" Then
        rs.MoveLast
        If rs.EOF = False Then
            If Format(rs.Fields("TglMasuk").Value, "yyyy/MM/dd") = Format(txttgl.Text, "yyyy/MM/dd") Then
                If IsNull(rs.Fields("TglPulang").Value) Then
                    GoTo hell
                End If
            End If
        End If
        'Else
        mstrIdPegawai = txtIDPegawai.Text
        If sp_Riwayat("A") = False Then Exit Sub
        If sp_simpan(txtIDPegawai.Text, txttgl.Text, txtkodestatus.Text, "A") = False Then Exit Sub
        Call txtidpegawai_Change
    End If
    Set rs = Nothing
hell:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim q As String

    If MSComm1.PortOpen = True Then
        q = MsgBox("Anda Yakin Mau Menutup Aplikasi Ini? Koneksi akan Terputus!", vbOKCancel, "Peringatan")
        If q = 2 Then
            Unload Me
            Cancel = 1
        Else
            Cancel = 0
        End If
    ElseIf MSComm1.PortOpen = False Then
        Unload Me
    End If

End Sub

Private Sub tmr_CekError_Timer()

    Set dbConn = Nothing
    Call conectServer

End Sub

Private Sub txtidpegawai_Change()
    On Error Resume Next
    txtNamaPegawai.Text = ""
    txtjk.Text = ""
    txtjabatan.Text = ""
    txttempattugas.Text = ""

    Set rs = Nothing
    strSQL = "SELECT * FROM v_AbsensiPegawai WHERE ID='" & txtIDPegawai.Text & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgAbsensi.DataSource = rs

    With dgAbsensi
        .Columns("IdPegawai").Value = txtIDPegawai.Text
        .Columns("Nama").Value = txtNamaPegawai.Text
        .Columns("Tgl. Masuk").Value = txttgl.Text
        .Columns("Tgl. Pulang").Value = txtakhir.Text
        .Columns("Tgl. Istirahat Awal").Value = txtistrahatawal.Text
        .Columns("Tgl. Istirahat Akhir").Value = txtistrahatakhir.Text
        .Columns("Status").Value = txtStatus.Text
        .Columns("Kode Status").Value = txtkodestatus.Text
        .Columns("NoRiwayat").Value = txtnoriwayat.Text
        .Columns("PIN").Value = txtPIN.Text
        .Columns("KdRuangan").Value = txtkode.Text

        .Columns("PIN").Width = 0
        .Columns("IdPegawai").Width = 0
        .Columns("NoRiwayat").Width = 0
        .Columns("Kode Status").Width = 0
        .Columns("Status").Width = 2600
        .Columns("JK").Width = 0
        .Columns("Ruangan").Width = 0
        .Columns("Shift").Width = 0
        .Columns("Jabatan").Width = 0
        .Columns("KdRuangan").Width = 0
        .Columns("Tgl. Istirahat Akhir").Width = 2000
        .Columns("Tgl. Istirahat Awal").Width = 2000
        .Columns("Tgl. Masuk").Width = 2000
        .Columns("Tgl. Pulang").Width = 2000
        .Columns("TglMulai").Width = 0

    End With
    With dgAbsensi
        txtIDPegawai.Text = .Columns("IdPegawai").Value
        txtNamaPegawai.Text = .Columns("Nama").Value
        txtjk.Text = .Columns("JK").Value
        txttempattugas.Text = .Columns("Ruangan")
        txtjabatan.Text = .Columns("Jabatan").Value
        txtkodestatus.Text = .Columns("Kode Status").Value
        txtnoriwayat.Text = .Columns("NoRiwayat").Value
        txtPIN.Text = .Columns("PIN")
        txtistrahatawal.Text = .Columns("Tgl. Istirahat Awal").Value
        txtistrahatakhir.Text = .Columns("Tgl. IStirahat Akhir").Value
        txtakhir.Text = .Columns("Tgl. Pulang").Value
        txtkode.Text = .Columns("KdRuangan").Value
    End With

End Sub

'timer untuk mengirim kode konfirmasi FRS
Private Sub kirim_konfirmasi_Timer()

    If fRS = add Then
        kirim_konfirmasi.Enabled = False
        If cek = 1 Then
            If add2 <> add Then
                MsgBox "Jumlah FRS-400 Tidak Sesuai dengan Alamat FRS-400 yang Dimasukkan !", vbOKOnly, "Pesan koneksi"
            End If
            Call reset1
            frmKFRS.cmdOk.Enabled = True 'diganti
            frmKFRS.cmdBatal.Enabled = True
            frmKFRS.Show
        ElseIf cek = 0 Then
            MsgBox "Tidak Ada Alat Yang Terhubung", vbOKOnly, "Pesan koneksi"
            Call Get_Disconnect
            cmdconnect.Enabled = True
            cmdsetting.Enabled = True
            cmdDisconnect.Enabled = False
            Unload frmKFRS
        End If
    Else
        fRS = (fRS + &H1)
        inbuff2 = ""
        MSComm1.Output = Chr$(&H2) & Chr$(&H5) & Chr$(fRS) & Chr$(&H4) & Chr$(&H3) & Chr$(fRS)
    End If

End Sub

'timer untuk mengirim kode untuk meminta hasil absensi
Private Sub minta_absensi_Timer()

    If c = 1 Then

        If Val(fRS) = Val(add2) Then
            fRS = &H0
            For i = 1 To add2
                caristring = indikator
                carichar = i
                indikator2 = InStr(caristring, carichar)
                If indikator2 <> 0 Then
                    List1.List(i) = "       " & "FRS - " & i & " (Beroperasi)"
                ElseIf indikator2 = 0 Then
                    List1.List(i) = "       " & "FRS - " & i & " (Tidak Beroperasi)"
                End If
            Next i
            indikator = ""
            indikator2 = ""
        End If

        code2 = code

        If fRS = &H0 Or Val(fRS) > 4 Or Val(add2) = 0 Then
            Call reset1
        End If

        fRS = (fRS + &H1)
        inbuff2 = ""
        MSComm1.Output = Chr$(&H2) & Chr$(&H5) & Chr$(fRS) & Chr$(&H1) & Chr$(&H3) & Chr$(code)

        'pemilihan kode akhir untuk minta absen
        Select Case add2
            Case 1
                code = &H4
            Case 2
                Select Case code
                    Case &H4
                        code = &H7
                    Case &H7
                        code = &H4
                End Select
            Case 3
                Select Case code
                    Case &H4
                        code = &H7
                    Case &H7
                        code = &H6
                    Case &H6
                        code = &H4
                End Select
        End Select

    ElseIf c = 2 Then

        If fRS = &H0 Or Val(fRS) > 4 Or Val(add2) = 0 Then
            Call reset1
            GoTo keluar
        End If

        inbuff2 = ""
        MSComm1.Output = Chr$(&H6) & Chr$(&H2) & Chr$(&H5) & Chr$(fRS) & Chr$(&H1) & Chr$(&H3) & Chr$(code2)

    End If

    If List1.List(1) = "       " & "FRS - 1" & "" And cekz = 2 Then
        If List1.List(2) = "       " & "FRS - 2" & "" And cekz = 2 Then
            If List1.List(3) = "       " & "FRS - 3" & "" And cekz = 2 Then
                Call reset1
                GoTo keluar
            End If
        End If
    End If

    If List1.List(1) = "       " & "FRS - 1 (Tidak Beroperasi)" Then
        If List1.List(2) = "       " & "FRS - 2 (Tidak Beroperasi)" Then
            If List1.List(3) = "       " & "FRS - 3 (Tidak Beroperasi)" Then
                Call reset1
                cekz = 2

                For i = 1 To add2
                    List1.List(i) = "       " & "FRS - " & i & ""
                Next i

                c = 1
                cekz = 2
                inbuff2 = ""

                GoTo keluar
            End If
        End If
    End If

keluar:
End Sub

Public Sub Get_Connect()

    On Error GoTo Handle_Error
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    MSComm1.PortOpen = True
    Exit Sub

    'pesan error
Handle_Error:
    MsgBox Error$, 48, "Konfirmasi Kesalahan Setting"
    Get_Disconnect

End Sub

Public Sub Get_Disconnect()

    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

    add2 = &H0
    Call reset1
    Call reset2

End Sub

Private Sub MSComm1_OnComm()
    Dim inbuff As Variant
    Dim partImg As String
    Dim part As String
    Dim mintaAbsen As String
    Dim keyData As Integer
    Dim strJumlahPIN As String
    Dim jumlahPINTersisa As Integer
    Dim pesan As VbMsgBoxResult

    Select Case MSComm1.CommEvent

            ' Errors
        Case comEventBreak                  'A Break was received.
        Case comEventCDTO                   'CD (RLSD) Timeout.
        Case comEventCTSTO                  'CTS Timeout.
        Case comEventDSRTO                  'DSR Timeout.
        Case comEventFrame                  'Framing Error.
        Case comEventOverrun                'Data Lost.
        Case comEventRxOver                 'Receive buffer overflow.
        Case comEventRxParity               'Parity Error.
        Case comEventTxFull                 'Transmit buffer full.
        Case comEventDCB                    'Unexpected error retrieving DCB

            ' Events
        Case comEvCD                        'Change in the CD line.
        Case comEvCTS                       'Change in the CTS line.
        Case comEvDSR                       'Change in the DSR line.
        Case comEvRing                      'Change in the Ring Indicator.
        Case comEvReceive                   'Received RThreshold # of chars.

            inbuff = MSComm1.Input
            inbuff2 = inbuff2 & inbuff
            If Len(inbuff2) > 3 Then
                indikator = indikator & Asc(Mid(inbuff2, 3, 1))
            End If

            If Len(inbuff2) = 30 Then
                c = 2
                Call absensi
                inbuff2 = ""
            ElseIf Len(inbuff2) = 6 Then
                c = 1
                inbuff2 = ""
            ElseIf Len(inbuff2) = 21 And Mid(inbuff2, 2, 1) = Chr(&H14) Then
                Call inisialisasi
                inbuff2 = ""
            ElseIf Len(inbuff2) = 15 And Mid(inbuff2, 2, 1) = Chr(&HE) Then
                Call HasilCek
                inbuff2 = ""
            ElseIf Len(inbuff2) = 215 And Mid(inbuff2, 2, 1) = Chr(&HD6) Then
                Dim protokolUpload As String

                If fp = 0 Then
                    FP1 = vbEmpty
                    FP2 = vbEmpty
                    FP3 = vbEmpty
                    FP4 = vbEmpty
                    gmbrFP = ""
                End If

                fp = fp + 1

                Select Case fp
                    Case 1
                        FP1 = Mid(inbuff2, 14, 200)
                        subSimpanImageFP pinSimpan, FP1, "FingerPrint1"
                    Case 2
                        FP2 = Mid(inbuff2, 14, 200)
                        subSimpanImageFP pinSimpan, FP2, "FingerPrint2"
                    Case 3
                        FP3 = Mid(inbuff2, 14, 200)
                        subSimpanImageFP pinSimpan, FP3, "FingerPrint3"
                    Case 4
                        FP4 = Mid(inbuff2, 14, 200)
                        subSimpanImageFP pinSimpan, FP4, "FingerPrint4"
                End Select
                inbuff2 = ""
                If fp = 1 Or fp = 2 Or fp = 3 Then
                    MSComm1.Output = Chr$(&H6)
                ElseIf fp = 4 Then
                    fp = 0
                    If Not bolFullUpload Then Exit Sub
balikmaning:
                    idxListViewPIN = idxListViewPIN + 1
                    If idxListViewPIN > frmPINAbsensiPegawai.ListView1.ListItems.Count Then
                        MsgBox "Upload semua PIN dari FRS-" & frsLihatPIN & " selesai!", vbInformation, "Perhatian"
                        idxListViewPIN = 0
                        bolFullUpload = False
                        frmAbsensiPegawai.minta_absensi.Enabled = True
                        frmAbsensiPegawai.tmr_CekError.Enabled = True
                        Unload frmStatusProses
                        Exit Sub
                    End If
                    frmStatusProses.pgbStatus.Value = idxListViewPIN
                    If Not frmPINAbsensiPegawai.ListView1.ListItems.Item(idxListViewPIN).Checked Then GoTo balikmaning
                    If frmPINAbsensiPegawai.ListView1.ListItems.Item(idxListViewPIN).SubItems(5) = "" Then GoTo balikmaning
                    pinSimpan = frmPINAbsensiPegawai.ListView1.ListItems.Item(idxListViewPIN).Text

                    strSQL = "SELECT FingerPrint4 FROM PINAbsensiPegawai WHERE PINAbsensi='" & pinSimpan & "'"
                    Set rs = New ADODB.recordset
                    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

                    If Not IsNull(rs.Fields.Item(0).Value) Then
                        rs.Close
                        GoTo balikmaning
                    End If
                    Dim crFRS As Integer
                    frsLihatPIN = frmPINAbsensiPegawai.txtCek.Text
                    crFRS = InStr(1, frmPINAbsensiPegawai.ListView1.ListItems.Item(idxListViewPIN).SubItems(1), frsLihatPIN, vbTextCompare)
                    If crFRS = 0 Then GoTo balikmaning
                    frmPINAbsensiPegawai.ListView1.ListItems.Item(idxListViewPIN).SubItems(2) = "Y"
                    protokolUpload = funcBuatProtokolUpload(frsLihatPIN, pinSimpan)
                    Me.MSComm1.Output = Chr$(&H6) & protokolUpload
                End If

            ElseIf inbuff2 = Chr$(&H6) And statDownload Then
                If dl = 1 Then
                    partImg = "FingerPrint2"
                ElseIf dl = 2 Then
                    partImg = "FingerPrint3"
                ElseIf dl = 3 Then
                    partImg = "FingerPrint4"
                End If
                dl = dl + 1
                part = dl
                If dl > 4 Then
                    dl = 0
                    statDownload = False
                    inbuff2 = ""
                    Exit Sub
                End If

                Dim protokolDownload As String
                Set rs = New ADODB.recordset
                rs.Open "SELECT " & partImg & " FROM PINAbsensiPegawai " & _
                "WHERE PINAbsensi=" & funcPrepareString(CInt(pinCek)), dbConn, 3, 3
                Dim strImageData As String
                strImageData = rs.Fields(0).Value
                varImageData = funcConvertHexKeImage(strImageData)
                protokolDownload = Chr$(&H2) & Chr$(&HD7) & Chr$(frsTujuan) & Chr$(&H10) & _
                pinCek & Chr$(&H30) & part & varImageData & Chr$(&H3)
                protokolDownload = protokolDownload & Chr$(cekSum(protokolDownload))
                rs.Close
                Me.MSComm1.Output = protokolDownload
                inbuff2 = ""
            ElseIf statusTransfer And ((Mid(inbuff2, 2, 1) = Chr$(&HD) And Len(inbuff2) = 14) Or (Mid(inbuff2, 3, 1) = Chr$(&HD) And Len(inbuff2) = 15)) Then
                Dim cr As Integer

                cr = InStr(1, inbuff2, pinCek, vbTextCompare)
                frmPINAbsensiPegawai.MousePointer = 0
                If cr <> 0 Then
                    MsgBox "Transfer data PIN & finger print untuk nomor PIN " & CInt(pinCek) & " ke FRS-" & frsTujuan & _
                    " berhasil.", vbInformation, "Perhatian.."
                Else
                    MsgBox "Transfer data PIN & finger print untuk nomor PIN " & CInt(pinCek) & " ke FRS-" & frsTujuan & _
                    " gagal!", vbExclamation, "Perhatian.."
                End If
                statusTransfer = False
                inbuff2 = ""
                frmAbsensiPegawai.minta_absensi.Enabled = True
                frmAbsensiPegawai.tmr_CekError.Enabled = True
            ElseIf statusPinHapus And Len(inbuff2) = 14 Then
                Dim cr1 As Integer

                cr1 = InStr(1, inbuff2, "FFFF", vbTextCompare)
                If cr1 = 0 Then
                    MsgBox "Hapus PIN: " & pinCek & " pada FRS-" & frsHapus & " berhasil.", vbInformation, "Hapus PIN"
                Else
                    MsgBox "Hapus PIN: " & pinCek & " pada FRS-" & frsHapus & " gagal!", vbExclamation, "Hapus PIN"
                End If
                statusPinHapus = False
                inbuff2 = ""
                frmAbsensiPegawai.minta_absensi.Enabled = True
            ElseIf bolCekJumlahPIN And Len(inbuff2) = 10 Then

                keyData = InStr(1, inbuff2, Chr$(&H13), vbTextCompare)
                If keyData <> 0 Then
                    strJumlahPIN = Mid(inbuff2, keyData + 1, 4)
                    jumlahTotalPIN = CInt(strJumlahPIN)
                    pesan = MsgBox("Lihat " & jumlahTotalPIN & " PIN di FRS-" & frsTujuan & "?", vbInformation Or vbYesNo, "Perhatian...")
                    If pesan = vbYes Then
                        frmStatusProses.Show
                        idxTempDataPIN = 0
                        nPrepareUpload = 0
                        Me.MSComm1.Output = funcPrepareFullUploadProtocol("0")
                        Me.minta_absensi.Enabled = False
                        Me.tmr_CekError.Enabled = False
                        bolCekJumlahPIN = False
                        bolPrepareFullUpload = True
                    Else
                        Me.minta_absensi.Enabled = True
                        Me.tmr_CekError.Enabled = True
                        bolCekJumlahPIN = False
                        bolPrepareFullUpload = False
                    End If
                End If
                inbuff2 = ""
            ElseIf bolPrepareFullUpload And cekSum(inbuff2) = "0" Then
                If cekSum(inbuff2) <> "0" Then
                    bolPrepareFullUpload = False
                    Exit Sub
                End If

                keyData = InStr(1, inbuff2, Chr$(&H13), vbTextCompare)
                If keyData <> 0 Then
                    strJumlahPIN = Mid(inbuff2, 5, 4)
                    jumlahPIN = CInt(strJumlahPIN)
                    If Len(inbuff2) < (jumlahPIN * 4) + 10 Then GoTo lompat
                    With frmStatusProses
                        subAmbilPIN inbuff2, .pgbStatus, .lblStatus, .lblPIN
                    End With
                    If jumlahPIN = 0 Then
                        frmStatusProses.Show
                        frmPINAbsensiPegawai.tmrCetak.Enabled = True
                        strStatusSekarang = "cetak"
                        bolPrepareFullUpload = False
                        inbuff2 = ""
                        Exit Sub
                    End If
                    nPrepareUpload = nPrepareUpload + jumlahPIN
                    Me.MSComm1.Output = funcPrepareFullUploadProtocol(CStr(nPrepareUpload))
                End If
                inbuff2 = ""
lompat:
            ElseIf Len(inbuff2) > 216 Then
                c = 1
                inbuff2 = ""
            ElseIf Len(inbuff2) = 1 Then
                Call reset1
                Call reset2
            End If

        Case comEvSend                      'There are SThreshold number of characters in the transmit buffer.
        Case comEvEOF                       'An EOF character was found in the

    End Select
End Sub

'perintah pada tabel hasil konfirmasi FRS
Private Sub inisialisasi()
    cek = 1
    add2 = (add2 + &H1)
    With frmKFRS.MSFlexGrid1
        .Rows = .Rows + 1
        .TextMatrix(a, 0) = "FRS - " & Hex(Asc(Mid(inbuff2, 3, 1)))
        .TextMatrix(a, 1) = Mid(inbuff2, 12, 2) + "/" + Mid(inbuff2, 10, 2) + "/" + Mid(inbuff2, 6, 4)
        .TextMatrix(a, 2) = Mid(inbuff2, 14, 2) + ":" + Mid(inbuff2, 16, 2) + ":" + Mid(inbuff2, 18, 2)
        a = a + 1
    End With
End Sub

Sub kosong()
    txtIDPegawai.Text = ""
    txtNamaPegawai.Text = ""
    txtjk.Text = ""
    txtjabatan.Text = ""
    txttgl.Text = ""
    txtFingerPrint.Text = ""
    txtStatus.Text = ""
    txttempattugas.Text = ""
    txtPIN.Text = ""
    txtnoriwayat.Text = ""
    txtistrahatawal.Text = ""
    txtistrahatakhir.Text = ""
    txtkodestatus.Text = ""
    txtkode.Text = ""
End Sub

'perintah pada tabel hasil absensi dari FRS
Private Sub absensi()
    Dim bnt As String
    Call kosong

    If Mid(inbuff2, 6, 1) = 1 Then
        txtPIN.Text = Format(Mid(inbuff2, 21, 8), "########")
        If txtIDPegawai.Text <> "" Then
            txtFingerPrint.Text = "FRS - " & Hex(Asc(Mid(inbuff2, 3, 1)))
            txttgl.Text = DateSerial(Mid(inbuff2, 7, 4), Mid(inbuff2, 11, 2), Mid(inbuff2, 13, 2)) & " " & Mid(inbuff2, 15, 2) & ":" & Mid(inbuff2, 17, 2) & ":" & Mid(inbuff2, 19, 2)
            If Mid(inbuff2, 5, 1) = 1 Then
                txtStatus.Text = "MASUK"
            ElseIf Mid(inbuff2, 5, 1) = 2 Then
                txtStatus.Text = "PULANG"
            ElseIf Mid(inbuff2, 5, 1) = 3 Then
                txtStatus.Text = "ISTIRAHAT AWAL"
            ElseIf Mid(inbuff2, 5, 1) = 4 Then
                txtStatus.Text = "ISTIRAHAT AKHIR"
            End If

        End If
    ElseIf Mid(inbuff2, 6, 1) = 0 Then
        Exit Sub
    End If
End Sub

Sub loadgrid()
    Set rs = Nothing
    strSQL = "SELECT * FROM v_P_Absensi"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set DataGrid1.DataSource = rs
End Sub

Private Sub txtPIN_Change()
    On Error Resume Next
    Set rs = Nothing
    strSQL = "SELECT * FROM v_P_Absensi WHERE PIN='" & txtPIN.Text & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then Exit Sub
    Set DataGrid1.DataSource = rs

    txtIDPegawai.Text = DataGrid1.Columns("ID").Value
    txtNamaPegawai.Text = DataGrid1.Columns("Nama").Value
    txtjk.Text = DataGrid1.Columns("JK").Value
    txttempattugas.Text = DataGrid1.Columns("Ruangan").Value
    txtjabatan.Text = DataGrid1.Columns("Jabatan").Value
End Sub

Private Function sp_Riwayat(f_Status) As Boolean
    On Error GoTo errLoad
    sp_Riwayat = True
    Set adoComm = New ADODB.Command
    With adoComm

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglRiwayat", adDate, adParamInput, , Format(Now, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, txtkode.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .Parameters.Append .CreateParameter("OutputNoRiwayat", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Riwayat"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_Riwayat = False
        Else
            If Not IsNull(.Parameters("OutputNoRiwayat").Value) Then txtnoriwayat.Text = .Parameters("OutputNoRiwayat").Value
        End If
        txtnoriwayat.Text = .Parameters("OutputNoRiwayat").Value
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
errLoad:
End Function

Private Sub pin()

    Dim Jml

    R = 9
    Jml = Val(Mid(inbuff2, 5, 4))
    With frmKPIN.MSFlexGrid1
        For i = 0 To (Jml - 1)
            pin1 = "00" & (Hex(Asc(Mid(inbuff2, R, 1))))
            pin1 = Right(pin1, 2)
            pin2 = "00" & (Hex(Asc(Mid(inbuff2, (R + 1), 1))))
            pin2 = Right(pin2, 2)
            pin3 = "00" & (Hex(Asc(Mid(inbuff2, (R + 2), 1))))
            pin3 = Right(pin3, 2)
            pin4 = "00" & (Hex(Asc(Mid(inbuff2, (R + 3), 1))))
            pin4 = Right(pin4, 2)
            pinz = Val(pin1 & pin2 & pin3 & pin4)

            strSQL = "SELECT  PINAbsensi FROM PINAbsensiPegawai  ORDER BY PINAbsensi"
            Call msubRecFO(rs, strSQL)
            Set dgPin.DataSource = rs
            a = a + 1
            .Rows = .Rows + 1

sama:

            R = R + 4

        Next i
    End With
End Sub

Private Sub txtstatus_Change()

    If txtStatus.Text = "PULANG" Then
        txtkodestatus.Text = "04"
        txtakhir.Text = txttgl.Text
        Call cmdpulang_Click
    End If

    If txtStatus.Text = "ISTIRAHAT AWAL" Then
        txtkodestatus.Text = "02"
        txtistrahatawal.Text = txttgl.Text
        Call cmdawal_Click
    End If

    If txtStatus.Text = "ISTIRAHAT AKHIR" Then
        txtkodestatus.Text = "03"
        txtistrahatakhir.Text = txttgl.Text
        Call cmdakhir_Click
    End If

    If txtStatus.Text = "MASUK" Then
        txtkodestatus.Text = "01"
        Call cmdmasuk_Click
    End If
End Sub

Private Sub cmdpulang_Click()
    On Error Resume Next
    Set rs = Nothing
    strSQL = "SELECT TglMasuk, TglPulang" _
    & " FROM AbsensiPegawai" _
    & " WHERE IdPegawai='" & txtIDPegawai.Text & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    If txtStatus.Text = "PULANG" Then
        rs.MoveLast
        If rs.EOF = False Then
            If Not IsNull(Format(rs.Fields("TglMasuk").Value, "yyyy/MM/dd")) Then
                If IsNull(rs.Fields("TglPulang").Value) Then
                    If UpdateAbsensi("A") = False Then Exit Sub
                    Call txtidpegawai_Change
                End If
            Else
                GoTo hell
            End If
        End If
    End If

    Exit Sub
    Set rs = Nothing
hell:
End Sub

Private Function sp_simpan(f_IdPegawai As String, F_TglPulang As String, f_kdstatusabsensi, f_Status As String) As Boolean
    Dim status As String
    sp_simpan = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, IIf(txtnoriwayat.Text = "", Null, txtnoriwayat.Text))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(txttgl.Text, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglIstirahatAwal", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglIstirahatAkhir", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("KdStatusAbsensi", adChar, adParamInput, 2, txtkodestatus.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .Parameters.Append .CreateParameter("outputno", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_ABSENSIPEGAWAI"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Pegawai", vbCritical, "Validasi"
            sp_simpan = False
        Else
            If Not IsNull(.Parameters("outputno").Value) Then txtnoriwayat = .Parameters("Outputno").Value
            mstrIdPegawai = txtIDPegawai.Text
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
End Function

Private Function UpdateAbsensi(f_Status As String) As Boolean
    UpdateAbsensi = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_Value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtnoriwayat.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIDPegawai.Text)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(txttgl.Text, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , IIf(txtakhir.Text = "", Null, Format(txtakhir.Text, "yyyy/MM/dd HH:mm:ss")))
        .Parameters.Append .CreateParameter("TglIstirahatAwal", adDate, adParamInput, , IIf(txtistrahatawal.Text = "", Null, Format(txtistrahatawal.Text, "yyyy/MM/dd HH:mm:ss")))
        .Parameters.Append .CreateParameter("TglIstirahatAkhir", adDate, adParamInput, , IIf(txtistrahatakhir.Text = "", Null, Format(txtistrahatakhir.Text, "yyyy/MM/dd HH:mm:ss")))
        .Parameters.Append .CreateParameter("kdtstatusabsensi", adChar, adParamInput, 2, txtkodestatus.Text)
        .ActiveConnection = dbConn
        .CommandText = "Update_AbsensiPegawaiPULANG"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            UpdateAbsensi = False
            MsgBox "Error", vbExclamation, "Validasi"
        Else
        End If
        Call deleteADOCommandParameters(dbcmd)
    End With
End Function

Private Function UpdateAbsensiAwal(f_Status As String) As Boolean
    UpdateAbsensiAwal = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_Value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtnoriwayat.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIDPegawai.Text)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(txttgl.Text, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglIstirahatAwal", adDate, adParamInput, , IIf(txtistrahatawal.Text = "", Null, Format(txtistrahatawal.Text, "yyyy/MM/dd HH:mm:ss")))
        .Parameters.Append .CreateParameter("TglIstirahatAkhir", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("kdtstatusabsensi", adChar, adParamInput, 2, txtkodestatus.Text)
        
        .ActiveConnection = dbConn
        .CommandText = "Update_AbsensiISTIRAHATAWAL"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            UpdateAbsensiAwal = False
            MsgBox "Error", vbExclamation, "Validasi"
        Else
        End If
        Call deleteADOCommandParameters(dbcmd)
    End With
End Function

Private Function UpdateAbsensiAkhir(f_Status As String) As Boolean
    UpdateAbsensiAkhir = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_Value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtnoriwayat.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIDPegawai.Text)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(txttgl.Text, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglIstirahatAwal", adDate, adParamInput, , Format(txtistrahatawal.Text, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglIstirahatAkhir", adDate, adParamInput, , IIf(txtistrahatakhir.Text = "", Null, Format(txtistrahatakhir.Text, "yyyy/MM/dd HH:mm:ss")))
        .Parameters.Append .CreateParameter("kdtstatusabsensi", adChar, adParamInput, 2, txtkodestatus.Text)
        
        .ActiveConnection = dbConn
        .CommandText = "Update_AbsensiISTIRAHATAKHIR"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            UpdateAbsensiAkhir = False
            MsgBox "Error", vbExclamation, "Validasi"
        Else
        End If
        Call deleteADOCommandParameters(dbcmd)
    End With
End Function

Sub cekmasuk()
    Set rs = Nothing
    strSQL = "SELECT Max(NoRiwayat), TglMasuk" _
    & " FROM AbsensiPegawai" _
    & " WHERE IdPegawai='" & txtIDPegawai.Text & "'" _
    & " GROUP BY TglMasuk"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set rs = Nothing
End Sub

Sub cekistirahatawal()
    Set rs = Nothing
    strSQL = "SELECT Max(NoRiwayat), TglIstirahatAwal" _
    & " FROM AbsensiPegawai" _
    & " WHERE IdPegawai='" & txtIDPegawai.Text & "'" _
    & " GROUP BY TglIstirahatAwal"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set rs = Nothing
End Sub

Sub cekistirahatakhir()
    Set rs = Nothing
    strSQL = "SELECT Max(NoRiwayat), TglIstirahatAkhir" _
    & " FROM AbsensiPegawai" _
    & " WHERE IdPegawai='" & txtIDPegawai.Text & "'" _
    & " GROUP BY TglIstirahatAkhir"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set rs = Nothing
End Sub

Sub cekpulang()
    Set rs = Nothing
    strSQL = "SELECT Max(NoRiwayat), TglPulang" _
    & " FROM AbsensiPegawai" _
    & " WHERE IdPegawai='" & txtIDPegawai.Text & "'" _
    & " GROUP BY TglPulang"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set rs = Nothing
End Sub

Sub fingerprint()

    If fp = 0 Then
        FP1 = ""
        FP2 = ""
        FP3 = ""
        FP4 = ""
        gmbrFP = ""
    End If

    fp = fp + 1

    Select Case fp
        Case 1
            FP1 = Mid(inbuff2, 14, 200)
            subSimpanImageFP pinSimpan, FP1, "FingerPrint1"
        Case 2
            FP2 = Mid(inbuff2, 14, 200)
            subSimpanImageFP pinSimpan, FP2, "FingerPrint2"
        Case 3
            FP3 = Mid(inbuff2, 14, 200)
            subSimpanImageFP pinSimpan, FP3, "FingerPrint3"
        Case 4
            FP4 = Mid(inbuff2, 14, 200)
            subSimpanImageFP pinSimpan, FP4, "FingerPrint4"
            gmbrFP = FP1 & FP2 & FP3 & FP4
    End Select
End Sub

Sub reset1()

    fRS = &H0
    code = &H4
    a = 0
    c = 1
    cekz = 1
    cek = 0
    inbuff2 = ""

End Sub

Sub reset2()

    hasil = "0000"
    jumlah = "0000"
    frsHapus = &H0
    simpanPIN = ""
    key = &H17

End Sub

Sub conectServer()

    On Error GoTo NoConnz
    With dbConn
        If .State = adStateOpen Then Exit Sub
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & strDatabaseName & ";Data Source=" & strServerName
        .Open
        If dbConn.State = adStateOpen Then
            If minta_absensi.Enabled = False Then
                minta_absensi.Enabled = True
            End If
        Else
            minta_absensi.Enabled = False
            List1.AddItem "JARINGAN SERVER TERGANGGU"
            subNetSend "simrs_04", "Ada gangguan jaringan di Aplikasi HRD. Mohon diperiksa."
            subNetSend "simrs_5", "Ada gangguan jaringan di Aplikasi HRD. Mohon diperiksa."
        End If
    End With
    Exit Sub

NoConnz:
    minta_absensi.Enabled = False
    List1.AddItem "JARINGAN SERVER TERGANGGU"
End Sub

Sub HasilCek()
    Dim ada As Integer

    ada = InStr(1, inbuff2, pinCek, vbTextCompare)
    If ada <> 0 Then
        MsgBox "No. PIN " & frmPINAbsensiPegawai.txtNoPIN.Text & " Ada Pada Alamat FRS -" & frmPINAbsensiPegawai.txtCek.Text, vbInformation, "Pesan Hasil Cek PIN"
    Else
        MsgBox "No. PIN " & frmPINAbsensiPegawai.txtNoPIN.Text & " Tidak Ada Pada Alamat FRS -" & frmPINAbsensiPegawai.txtCek.Text, vbExclamation, "Pesan Hasil Cek PIN"
    End If
    Me.minta_absensi.Enabled = True
End Sub

Private Sub subSimpanImageFP(ByVal refPIN As String, ByVal imageData As Variant, ByVal fieldFP As String)

    Dim strImageData As String

    Set adoComm = New ADODB.Command
    strImageData = funcConvertImageKeHex(imageData)

    adoComm.ActiveConnection = dbConn
    adoComm.CommandText = "UPDATE PINAbsensiPegawai SET " & _
    fieldFP & "='" & strImageData & "'" & _
    " WHERE PINAbsensi=" & funcPrepareString(refPIN)
    adoComm.CommandType = adCmdText
    adoComm.Execute

    Exit Sub
salah:
    MsgBox "salah"
End Sub
