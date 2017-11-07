VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMasterAbsensi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Master Absensi"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "frmMasterAbsensi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   6285
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
      Left            =   2640
      TabIndex        =   5
      Top             =   6480
      Width           =   1095
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
      Left            =   3840
      TabIndex        =   4
      Top             =   6480
      Width           =   1095
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
      Left            =   5040
      TabIndex        =   3
      Top             =   6480
      Width           =   1095
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
      Left            =   1440
      TabIndex        =   2
      Top             =   6480
      Width           =   1095
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
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Jenis Hari"
      TabPicture(0)   =   "frmMasterAbsensi.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(2)=   "dgJenisHari"
      Tab(0).Control(3)=   "txtnamajenishari"
      Tab(0).Control(4)=   "txtkdjenishari"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Hari Libur"
      TabPicture(1)   =   "frmMasterAbsensi.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkTgl"
      Tab(1).Control(1)=   "txtkdharilibur"
      Tab(1).Control(2)=   "txtnamaharilibur"
      Tab(1).Control(3)=   "txtketerangan"
      Tab(1).Control(4)=   "dcJnsharilibur"
      Tab(1).Control(5)=   "dgHarilibur"
      Tab(1).Control(6)=   "dtpTglLibur"
      Tab(1).Control(7)=   "Label13"
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(9)=   "Label6"
      Tab(1).Control(10)=   "Label2"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Jam Kerja"
      TabPicture(2)   =   "frmMasterAbsensi.frx":0D02
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label11"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label14"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "dcNamaShift"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtjampulang"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtistrahatakhir"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtistrahatawal"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtJammasuk"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "dgshift"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtKdshift"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtnamashift"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Shift Kerja"
      TabPicture(3)   =   "frmMasterAbsensi.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtJenisShift2"
      Tab(3).Control(1)=   "txtKdShiftKerja2"
      Tab(3).Control(2)=   "dgShift2"
      Tab(3).Control(3)=   "Label12"
      Tab(3).Control(4)=   "Label1"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Konversi Shift Karyawan"
      TabPicture(4)   =   "frmMasterAbsensi.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "dcTempatBertugas"
      Tab(4).Control(1)=   "lvjadwalkerja"
      Tab(4).Control(2)=   "dcShiftKerja2"
      Tab(4).Control(3)=   "dcShiftKerja"
      Tab(4).Control(4)=   "Label16"
      Tab(4).Control(5)=   "Label15"
      Tab(4).ControlCount=   6
      Begin VB.TextBox txtJenisShift2 
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
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   32
         Top             =   1380
         Width           =   3615
      End
      Begin VB.TextBox txtKdShiftKerja2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         Left            =   -74640
         MaxLength       =   2
         TabIndex        =   31
         Top             =   1380
         Width           =   855
      End
      Begin VB.CheckBox chkTgl 
         Caption         =   "Tgl Hari Libur"
         Height          =   255
         Left            =   -71280
         TabIndex        =   30
         Top             =   1020
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtnamashift 
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
         Height          =   375
         Left            =   480
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtKdshift 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         Left            =   480
         MaxLength       =   3
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtkdharilibur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         Left            =   -74520
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtnamaharilibur 
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
         Left            =   -74520
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1860
         Width           =   1695
      End
      Begin VB.TextBox txtketerangan 
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
         Left            =   -72600
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1860
         Width           =   3135
      End
      Begin VB.TextBox txtkdjenishari 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         Left            =   -74520
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtnamajenishari 
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
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1380
         Width           =   3615
      End
      Begin MSDataGridLib.DataGrid dgJenisHari 
         Height          =   3135
         Left            =   -74520
         TabIndex        =   8
         Top             =   1860
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5530
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcJnsharilibur 
         Height          =   315
         Left            =   -73440
         TabIndex        =   14
         Top             =   1260
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSDataGridLib.DataGrid dgHarilibur 
         Height          =   2655
         Left            =   -74520
         TabIndex        =   15
         Top             =   2340
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4683
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpTglLibur 
         Height          =   315
         Left            =   -71280
         TabIndex        =   16
         Top             =   1260
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109510656
         UpDown          =   -1  'True
         CurrentDate     =   39282
      End
      Begin MSDataGridLib.DataGrid dgshift 
         Height          =   2415
         Left            =   480
         TabIndex        =   23
         Top             =   2580
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgShift2 
         Height          =   3135
         Left            =   -74640
         TabIndex        =   33
         Top             =   1860
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5530
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcTempatBertugas 
         Height          =   330
         Left            =   -74760
         TabIndex        =   37
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
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
      Begin MSComctlLib.ListView lvjadwalkerja 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   38
         Top             =   1200
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4895
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Diagnosa"
            Object.Width           =   13229
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcShiftKerja2 
         Height          =   315
         Left            =   -71640
         TabIndex        =   39
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSDataListLib.DataCombo dcShiftKerja 
         Height          =   315
         Left            =   -74760
         TabIndex        =   41
         Top             =   3360
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSMask.MaskEdBox txtJammasuk 
         Height          =   390
         Left            =   2040
         TabIndex        =   42
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtistrahatawal 
         Height          =   390
         Left            =   3840
         TabIndex        =   43
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtistrahatakhir 
         Height          =   390
         Left            =   2040
         TabIndex        =   44
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtjampulang 
         Height          =   390
         Left            =   3840
         TabIndex        =   45
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcNamaShift 
         Height          =   315
         Left            =   480
         TabIndex        =   46
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Shift Kerja"
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
         Left            =   -71640
         TabIndex        =   40
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Tempat Bertugas"
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
         Left            =   -74760
         TabIndex        =   36
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Shift Kerja"
         Height          =   195
         Left            =   -73560
         TabIndex        =   35
         Top             =   1140
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   195
         Left            =   -74640
         TabIndex        =   34
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Jam Istirahat Akhir"
         Height          =   195
         Left            =   2040
         TabIndex        =   29
         Top             =   1620
         Width           =   1290
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Jam Istirahat Awal"
         Height          =   195
         Left            =   3840
         TabIndex        =   28
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jam Pulang"
         Height          =   195
         Left            =   3840
         TabIndex        =   27
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Shift"
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   1620
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jam Masuk"
         Height          =   195
         Left            =   2040
         TabIndex        =   25
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   195
         Left            =   480
         TabIndex        =   24
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   195
         Left            =   -74520
         TabIndex        =   20
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Hari"
         Height          =   195
         Left            =   -73440
         TabIndex        =   19
         Top             =   1020
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nama Hari Libur"
         Height          =   195
         Left            =   -74520
         TabIndex        =   18
         Top             =   1620
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   195
         Left            =   -72600
         TabIndex        =   17
         Top             =   1620
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   195
         Left            =   -74520
         TabIndex        =   10
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Jenis Hari"
         Height          =   195
         Left            =   -73440
         TabIndex        =   9
         Top             =   1140
         Width           =   1155
      End
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4680
      Picture         =   "frmMasterAbsensi.frx":0D56
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterAbsensi.frx":1ADE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterAbsensi.frx":449F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMasterAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strKey As String
Dim i, j As Integer

Sub subDCSource()
    On Error Resume Next
    strSQL = "SELECT * FROM JenisHari order by JenisHari"
    Call msubDcSource(dcJnsharilibur, rs, strSQL)

    strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan order by NamaRuangan"
    Call msubDcSource(dcTempatBertugas, rs, strSQL)

    strSQL = "Select IdShift, Dinaskerja From DinasKerja"
    Call msubDcSource(dcShiftKerja2, rs, strSQL)
    
    strSQL = "select IdShift,DinasKerja from DinasKerja"
    Call msubDcSource(dcNamaShift, rs, strSQL)
End Sub

Sub sp_simpan()
    Select Case sstDataPenunjang.Tab
        Case 0 'jenis hari
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdJenisHari", adChar, adParamInput, 2, txtkdjenishari.Text)
                .Parameters.Append .CreateParameter("JenisHari", adVarChar, adParamInput, 20, txtnamajenishari.Text)
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Null)
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 20, Null)
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , 1)
                .Parameters.Append .CreateParameter("OutputKdJenisHari", adChar, adParamOutput, 2, Null)

                .ActiveConnection = dbConn
                .CommandText = "AU_JenisHari"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"
                Else
                    If Not IsNull(.Parameters("OutputKdJenisHari").Value) Then txtkdjenishari = .Parameters("OutputKdjenishari").Value
                End If
                Call deleteADOCommandParameters(dbcmd)
            End With
            cmdBatal_Click

        Case 1 'Hari Libur
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdHariLibur", adVarChar, adParamInput, 3, txtkdharilibur.Text)
                If chkTgl.Value = 1 Then
                    .Parameters.Append .CreateParameter("TglHariLibur", adDate, adParamInput, , Format(dtpTglLibur.Value, "yyyy/MM/dd"))
                Else
                    .Parameters.Append .CreateParameter("TglHariLibur", adDate, adParamInput, , Null)
                End If
                .Parameters.Append .CreateParameter("NamaHariLibur", adVarChar, adParamInput, 50, txtnamaharilibur.Text)
                If txtKeterangan.Text <> "" Then
                    .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, IIf(txtKeterangan.Text = "", Null, txtKeterangan.Text))
                Else
                    .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
                End If
                .Parameters.Append .CreateParameter("KdJenisHari", adChar, adParamInput, 2, Trim(dcJnsharilibur.BoundText))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Null)
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Null)
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , 1)
                .Parameters.Append .CreateParameter("OutputKdharilibur", adVarChar, adParamOutput, 3, Null)

                .ActiveConnection = dbConn
                .CommandText = "AU_HariLibur"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"
                Else
                    If Not IsNull(.Parameters("OutputKdHariLibur").Value) Then txtkdharilibur = .Parameters("OutputKdharilibur").Value
                End If
                Call deleteADOCommandParameters(dbcmd)
            End With
            cmdBatal_Click

        Case 2 'Jam kerja
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdShift", adVarChar, adParamInput, 2, Trim(txtKdshift))
'                .Parameters.Append .CreateParameter("NamaShift", adVarChar, adParamInput, 20, Trim(txtnamashift))
                .Parameters.Append .CreateParameter("NamaShift", adVarChar, adParamInput, 20, Trim(dcNamaShift.Text))
                .Parameters.Append .CreateParameter("JamMasuk", adChar, adParamInput, 5, Trim(txtJammasuk))
                .Parameters.Append .CreateParameter("JamPulang", adChar, adParamInput, 5, Trim(txtjampulang))
                .Parameters.Append .CreateParameter("JamIstirahatAwal", adChar, adParamInput, 5, Trim(txtistrahatawal))
                .Parameters.Append .CreateParameter("JamIstirahatAkhir", adChar, adParamInput, 5, Trim(txtistrahatakhir))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Null)
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Null)
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , 1)
                .Parameters.Append .CreateParameter("OutputKdShift", adVarChar, adParamOutput, 2, Null)

                .ActiveConnection = dbConn
                '.CommandText = "AU_ShiftKerja"
                .CommandText = "AU_ShiftKerja_New" '//yayang.agus 2014-08-08
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"
                Else
                    If Not IsNull(.Parameters("OutputKdShift").Value) Then txtKdshift = .Parameters("OutputKdShift").Value
                End If
                Call deleteADOCommandParameters(dbcmd)
            End With
            cmdBatal_Click

        Case 3 'Shift Kerja
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("IdShift", adChar, adParamInput, 2, txtKdShiftKerja2.Text)
                .Parameters.Append .CreateParameter("DinasKerja", adVarChar, adParamInput, 10, txtJenisShift2.Text)
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Null)
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Null)
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , 1)
                .Parameters.Append .CreateParameter("OutputIdShift", adChar, adParamOutput, 2, Null)

                .ActiveConnection = dbConn
                .CommandText = "AU_ShiftKerjaKaryawan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"
                Else
                    If Not IsNull(.Parameters("OutputIdShift").Value) Then txtKdShiftKerja2 = .Parameters("OutputIdShift").Value
                End If
                Call deleteADOCommandParameters(dbcmd)
            End With
            cmdBatal_Click

        Case 4 'Konversi

            For i = 1 To lvjadwalkerja.ListItems.Count
                If lvjadwalkerja.ListItems(i).Checked = True Then
                    If sp_KaryawanToShift(Right(lvjadwalkerja.ListItems(i).key, 10), "A") = False Then Exit Sub
                Else
                    If sp_KaryawanToShift(Right(lvjadwalkerja.ListItems(i).key, 10), "D") = False Then Exit Sub
                End If
            Next i

            Call Add_HistoryLoginActivity("AUD_KaryawanShiftNonShift")
            Call dcShiftKerja2_Change
            cmdBatal_Click

    End Select
End Sub

Public Function sp_KaryawanToShift(f_IdPegawai As String, f_status As String) As Boolean
    sp_KaryawanToShift = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_IdPegawai)
        .Parameters.Append .CreateParameter("IdShift", adChar, adParamInput, 2, dcShiftKerja2.BoundText)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KaryawanShiftNonShift"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data", vbCritical, "Validasi"
            sp_KaryawanToShift = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub chkTgl_Click()
    If chkTgl.Value = 1 Then
        dtpTglLibur.Enabled = True
    Else
        dtpTglLibur.Enabled = False
    End If
End Sub

Private Sub cmdBatal_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 'jenishari
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 1 'Hari Libur
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 2 'Jam kerja
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 3 'Shift Kerja
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 4 'Konversi
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True

    End Select
End Sub

Private Sub cmdHapus_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 'jenis hari
            Set rs = Nothing
            strSQL = "SELECT distinct dbo.jenishari.KdJenisHari  FROM dbo.HariLibur LEFT OUTER JOIN dbo.JenisHari ON dbo.HariLibur.KdJenisHari = dbo.JenisHari.KdJenisHari " & _
                     "where dbo.HariLibur.KdJenisHari = '" & dgJenisHari.Columns("Kode") & "' "
            Call msubRecFO(rsb, strSQL)
            If rsb.EOF = False Then
                MsgBox "Jenis hari sudah di pakai di data hari libur", vbInformation
            Exit Sub
            End If
            
            If MsgBox("Yakin akan menghapus jenis hari ?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
            strSQL = "delete JenisHari where KdJenisHari = '" & txtkdjenishari & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            MsgBox "Data berhasil dihapus", vbInformation
            Set rs = Nothing

        Case 1 'Hari Libur
            Set rs = Nothing
            
            If MsgBox("Yakin akan menghapus hari libur ?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
            '//YnA 2014-08-08
            Set rs = Nothing
            strSQL = "SELECT * FROM Datatanggal WHERE kdharilibur='" & txtkdharilibur & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            If rs.EOF = False Then
                MsgBox "Data masih terpakai, hapus terlebih dahulu di Hari Libur.", vbInformation
                Exit Sub
            End If
            Set rs = Nothing
            '//
            strSQL = "delete HariLibur where KdHariLibur = '" & txtkdharilibur & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            MsgBox "Data berhasil dihapus", vbInformation
            Set rs = Nothing

        Case 2 'Jam kerja
            Set rs = Nothing
            
            If MsgBox("Yakin akan menghapus Jam Kerja ?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
            strSQL = "delete shiftkerja_New where kdshift = '" & txtKdshift & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            MsgBox "Data berhasil dihapus", vbInformation
            Set rs = Nothing
        Case 3 ' Shift Kerja
            Set rs = Nothing
            
            If MsgBox("Yakin akan menghapus shift kerja ?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
            strSQL = "delete DinasKerja where IdShift = '" & txtKdShiftKerja2 & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockBatchOptimistic
            MsgBox "Data berhasil dihapus", vbInformation
            Set rs = Nothing
    End Select
    Call subLoadGridSource
    Call subKosong
End Sub

Private Sub cmdSimpan_Click()
    'On Error GoTo errSimpan
    Select Case sstDataPenunjang.Tab
        Case 0 'jenis hari
            If Periksa("text", txtnamajenishari, "Jenis Hari Harus diisi!!") = False Then Exit Sub
            Call sp_simpan
        Case 1  ' Hari Libur
            If Periksa("datacombo", dcJnsharilibur, "Harap di Isi Jenis Hari!") = False Then Exit Sub
            Call sp_simpan
        
        Case 2 'Jam kerja
'            If Periksa("text", txtnamashift, "Jam Kerja Harus diisi!!") = False Then Exit Sub
            If Periksa("datacombo", dcNamaShift, "Jam Kerja Harus diisi!!") = False Then Exit Sub
            If txtJammasuk.Text = "__:__" Then MsgBox "Jam Masuk Harus diisi!!", vbCritical: txtJammasuk.SetFocus: Exit Sub
            If txtistrahatawal.Text = "__:__" Then MsgBox "Jam Istirahat awal Harus diisi!!", vbCritical: txtistrahatawal.SetFocus: Exit Sub
            If txtistrahatakhir.Text = "__:__" Then MsgBox "Jam Istirahat akhir Harus diisi!!", vbCritical: txtistrahatakhir.SetFocus: Exit Sub
            If txtjampulang.Text = "__:__" Then MsgBox "Jam Pulang Harus diisi!!", vbCritical: txtjampulang.SetFocus: Exit Sub
           
            Call sp_simpan
        
        Case 3 'Shift Kerja
            If Periksa("text", txtJenisShift2, "Nama Shift Harus diisi!!") = False Then Exit Sub
            Call sp_simpan
        Case 4 'Konversi
            'If Periksa("datacombo", dcTempatBertugas, "Nama Tempat Bertugas harus diisi!") = False Then Exit Sub
            If Periksa("datacombo", dcShiftKerja2, "Nama Shift Bertugas harus diisi!") = False Then Exit Sub
            Call sp_simpan
            dcShiftKerja2.Text = ""
            MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
    End Select
    Call subLoadGridSource
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJnsharilibur_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then dtpTglLibur.SetFocus

On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcJnsharilibur.Text)) = 0 Then dtpTglLibur.SetFocus: Exit Sub
        If dcJnsharilibur.MatchedWithList = True Then dtpTglLibur.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdJenisHari, JenisHari FROM JenisHari WHERE JenisHari LIKE '%" & dcJnsharilibur.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcJnsharilibur.BoundText = dbRst(0).Value
        dcJnsharilibur.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNamaShift_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcNamaShift.Text)) = 0 Then txtistrahatakhir.SetFocus: Exit Sub
        If dcNamaShift.MatchedWithList = True Then txtistrahatakhir.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select IdShift,DinasKerja from DinasKerja WHERE DinasKerja LIKE '%" & dcNamaShift.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcNamaShift.BoundText = dbRst(0).Value
        dcNamaShift.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcShiftKerja2_Change()
    On Error GoTo hell
    
    Call loadListViewSource
    
    Call subLoadLView

    If dcShiftKerja2.MatchedWithList = False Then Exit Sub

    strSQL = "SELECT IdPegawai From ConvertIdPegawaiToShift WHERE (IdShift = '" & dcShiftKerja2.BoundText & "')"
    Call msubRecFO(rs, strSQL)
    For j = 0 To rs.RecordCount - 1
        For i = 1 To lvjadwalkerja.ListItems.Count
            If Right(lvjadwalkerja.ListItems(i).key, 10) = rs(0).Value Then
                lvjadwalkerja.ListItems.Item(i).Bold = True
                lvjadwalkerja.ListItems.Item(i).ForeColor = vbBlue
                lvjadwalkerja.ListItems.Item(i).Checked = True
            End If
        Next i
        rs.MoveNext
    Next j
    Exit Sub
hell:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub subLoadLView()
    On Error GoTo hell

    strSQL = "SELECT DISTINCT IdPegawai, NamaLengkap From v_TempatBertugas WHERE NamaRuangan = '" & dcTempatBertugas & "' ORDER BY IdPegawai"

    Call msubRecFO(rs, strSQL)
    'lvjadwalkerja.ListItems.clear
    Do While rs.EOF = False
        strKey = "key" & rs(0).Value
        lvjadwalkerja.ListItems.add , strKey, rs(1).Value
        rs.MoveNext
    Loop
    lvjadwalkerja.View = lvwList
    Exit Sub
hell:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub dcShiftKerja2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then cmdSimpan.SetFocus

On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcShiftKerja2.Text)) = 0 Then cmdSimpan.SetFocus: Exit Sub
        If dcTempatBertugas.MatchedWithList = True Then cmdSimpan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "Select IdShift, Dinaskerja From DinasKerja WHERE DinasKerja LIKE '%" & dcShiftKerja2.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcShiftKerja2.BoundText = dbRst(0).Value
        dcShiftKerja2.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcTempatBertugas_Change()

    Call loadListViewSource
End Sub

Public Sub loadListViewSource()
    On Error GoTo tangani

    strSQL = "SELECT DataCurrentPegawai.IdPegawai, DataPegawai.NamaLengkap " & _
    "FROM DataCurrentPegawai INNER JOIN " & _
    "DataPegawai ON DataCurrentPegawai.IdPegawai = DataPegawai.IdPegawai " & _
    "WHERE KdRuanganKerja = '" & dcTempatBertugas.BoundText & "' and KdStatus = '01' ORDER BY IdPegawai"
    Call msubRecFO(rs, strSQL)
    lvjadwalkerja.ListItems.clear
    While Not rs.EOF
        lvjadwalkerja.ListItems.add , "A" & rs(0).Value, rs(1).Value
        rs.MoveNext
    Wend
    lvjadwalkerja.Sorted = True

    If rs.RecordCount = 0 Then Exit Sub

    strSQL = "SELECT ID from v_JadwalKerjaNew WHERE KdShift = '" & dcShiftKerja.BoundText & "' AND KdRuangan = '" & dcTempatBertugas.BoundText & "' "

    Call msubRecFO(rs, strSQL)
    Do While rs.EOF = False
        lvjadwalkerja.ListItems("A" & rs(0)).Checked = True
        lvjadwalkerja.ListItems("A" & rs(0)).ForeColor = vbBlue
        lvjadwalkerja.ListItems("A" & rs(0)).Bold = True
        rs.MoveNext
    Loop
    Exit Sub

tangani:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub dcTempatBertugas_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcShiftKerja2.SetFocus

On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcTempatBertugas.Text)) = 0 Then dcShiftKerja2.SetFocus: Exit Sub
        If dcTempatBertugas.MatchedWithList = True Then dcShiftKerja2.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE NamaRuangan LIKE '%" & dcTempatBertugas.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcTempatBertugas.BoundText = dbRst(0).Value
        dcTempatBertugas.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgjenishari_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtkdjenishari.Text = dgJenisHari.Columns(0).Value
    txtnamajenishari.Text = dgJenisHari.Columns(1).Value
End Sub

Private Sub dgShift2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtKdShiftKerja2.Text = dgShift2.Columns(0).Value
    txtJenisShift2.Text = dgShift2.Columns(1).Value
End Sub

Private Sub dgHariLibur_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtkdharilibur.Text = dgHarilibur.Columns(0).Value
    dtpTglLibur.Value = dgHarilibur.Columns(1).Value
    txtnamaharilibur.Text = dgHarilibur.Columns(2).Value
    txtKeterangan.Text = dgHarilibur.Columns(3).Value
    dcJnsharilibur.Text = dgHarilibur.Columns(4).Value
End Sub

Private Sub dgshift_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdshift.Text = dgshift.Columns(0).Value
    txtnamashift.Text = dgshift.Columns(1).Value
    dcNamaShift.Text = dgshift.Columns(1).Value
    txtJammasuk = dgshift.Columns(2).Value
    txtjampulang.Text = dgshift.Columns(3).Value
    txtistrahatawal.Text = dgshift.Columns(4).Value
    txtistrahatakhir.Text = dgshift.Columns(5).Value
End Sub

Private Sub dtpTglLibur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtnamaharilibur.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subDCSource
    sstDataPenunjang.Tab = 0
    Call subLoadGridSource
End Sub

Sub subKosong()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0 'jenishari
            txtkdjenishari.Text = ""
            txtnamajenishari.Text = ""
        Case 1 'HariLibur
            txtkdharilibur.Text = ""
            dcJnsharilibur.Text = ""
            dtpTglLibur.Value = Now
            txtnamaharilibur.Text = ""
            txtKeterangan.Text = ""
        Case 2 'Jam Kerja
            txtKdshift.Text = ""
            txtnamashift.Text = ""
            dcNamaShift.Text = ""
            txtJammasuk.Text = "__:__"
            txtjampulang.Text = "__:__"
            txtistrahatawal.Text = "__:__"
            txtistrahatakhir.Text = "__:__"
        Case 3 'Shift Kerja
            txtKdShiftKerja2.Text = ""
            txtJenisShift2.Text = ""
        Case 4 ' Konversi
            dcTempatBertugas.Text = ""
            dcShiftKerja2.Text = ""
            lvjadwalkerja.ListItems.clear
    End Select
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call subDCSource
    Call subLoadGridSource
    Call cmdBatal_Click
    dtpTglLibur.Value = Now
End Sub

Sub subLoadGridSource()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0 ' jenishari
            Set rs = Nothing
            strSQL = "select * from jenishari"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJenisHari.DataSource = rs
            dgJenisHari.Columns(0).DataField = rs(0).Name
            dgJenisHari.Columns(1).DataField = rs(1).Name
            dgJenisHari.Columns(0).Caption = "Kode"
            dgJenisHari.Columns(1).Caption = "Jenis Hari"
            Set rs = Nothing

            txtnamajenishari.SetFocus

        Case 1  'Hari Libur
            Set rs = Nothing
            strSQL = "SELECT dbo.HariLibur.KdHariLibur, dbo.HariLibur.TglHariLibur, dbo.HariLibur.NamaHariLibur, dbo.HariLibur.Keterangan, dbo.JenisHari.JenisHari" & _
            " FROM dbo.HariLibur LEFT OUTER JOIN" & _
            " dbo.JenisHari ON dbo.HariLibur.KdJenisHari = dbo.JenisHari.KdJenisHari"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgHarilibur.DataSource = rs
            dgHarilibur.Columns(0).DataField = rs(0).Name
            dgHarilibur.Columns(1).DataField = rs(1).Name
            dgHarilibur.Columns(2).DataField = rs(2).Name
            dgHarilibur.Columns(3).DataField = rs(3).Name
            dgHarilibur.Columns(4).DataField = rs(4).Name

            dgHarilibur.Columns(0).Caption = "Kode"
            dgHarilibur.Columns(1).Caption = "Tgl. Libur"
            dgHarilibur.Columns(2).Caption = "Nama Hari Libur"
            dgHarilibur.Columns(3).Caption = "Keterangan"
            dgHarilibur.Columns(4).Caption = "Jenis Hari"
            dtpTglLibur.DataField = Now
            Set rs = Nothing

            dcJnsharilibur.SetFocus

        Case 2 'Jam kerja
            Set rs = Nothing
            strSQL = "select * from ShiftKerja_New"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgshift.DataSource = rs
            dgshift.Columns(0).DataField = rs(0).Name
            dgshift.Columns(1).DataField = rs(1).Name
            dgshift.Columns(2).DataField = rs(2).Name
            dgshift.Columns(3).DataField = rs(3).Name
            dgshift.Columns(4).DataField = rs(4).Name
            dgshift.Columns(5).DataField = rs(5).Name
            dgshift.Columns(0).Caption = "Kode"
            dgshift.Columns(0).Width = 750
            dgshift.Columns(1).Caption = "Shift"
            dgshift.Columns(1).Width = 1000
            dgshift.Columns(2).Caption = "Jam Masuk"
            dgshift.Columns(2).Width = 1500
            dgshift.Columns(3).Caption = "Jam Pulang"
            dgshift.Columns(3).Width = 1500
            dgshift.Columns(4).Caption = "Istirahat Awal"
            dgshift.Columns(4).Width = 1500
            dgshift.Columns(5).Caption = "Istirahat Akhir"
            dgshift.Columns(5).Width = 1500
            Set rs = Nothing

            txtJammasuk.SetFocus

        Case 3 ' Shift Kerja

            Set rs = Nothing
            strSQL = "select * from DinasKerja"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgShift2.DataSource = rs
            dgShift2.Columns(0).DataField = rs(0).Name
            dgShift2.Columns(1).DataField = rs(1).Name
            dgShift2.Columns(0).Caption = "Kode"
            dgShift2.Columns(1).Caption = "Shift Kerja"
            Set rs = Nothing

            txtJenisShift2.SetFocus

        Case 4 'Konversi Shift Kerja

            Call subDCSource

    End Select
End Sub

Private Sub txtistrahatakhir_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then If funcCekValidasiTgl("Jam", txtistrahatakhir) = "NoErr" Then txtjampulang.SetFocus
End Sub

Private Sub txtistrahatawal_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then If funcCekValidasiTgl("Jam", txtistrahatawal) = "NoErr" Then dcNamaShift.SetFocus
End Sub

Private Sub txtJamMasuk_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If funcCekValidasiTgl("Jam", txtJammasuk) = "NoErr" Then txtistrahatawal.SetFocus
    End If
End Sub

Private Sub txtJamPulang_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then If funcCekValidasiTgl("Jam", txtjampulang) = "NoErr" Then cmdSimpan.SetFocus
    
End Sub

Private Sub txtJenisShift2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtnamaharilibur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtnamajenishari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtnamashift_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtistrahatakhir.SetFocus
End Sub
