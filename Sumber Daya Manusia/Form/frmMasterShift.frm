VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterShift 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Shift Kerja"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "frmMasterShift.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   9015
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
      Left            =   6480
      TabIndex        =   8
      Top             =   8040
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
      Left            =   5280
      TabIndex        =   9
      Top             =   8040
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
      Left            =   7680
      TabIndex        =   10
      Top             =   8040
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
      Left            =   4080
      TabIndex        =   7
      Top             =   8040
      Width           =   1095
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   6735
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11880
      _Version        =   393216
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
      TabCaption(0)   =   "Shift"
      TabPicture(0)   =   "frmMasterShift.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dgshift"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtnamashift"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtKdshift"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtJamMasuk"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtJamPulang"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtJamMasukAwal"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtJamMasukAkhir"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtJamPulangAwal"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtJamPulangAkhir"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Konversi Shift To Ruangan"
      TabPicture(1)   =   "frmMasterShift.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Konversi Shift To Pegawai"
      TabPicture(2)   =   "frmMasterShift.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   -74400
         TabIndex        =   32
         Top             =   600
         Width           =   7575
         Begin VB.TextBox txtPegawai 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   3375
         End
         Begin VB.Frame fraPegawai 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Visible         =   0   'False
            Width           =   3495
            Begin MSDataGridLib.DataGrid dgPegawai 
               Height          =   3495
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   6165
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
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
         End
         Begin VB.TextBox txtIdPegawai 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   720
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   3375
         End
         Begin MSDataListLib.DataCombo dcShift2 
            Height          =   315
            Left            =   3720
            TabIndex        =   37
            Top             =   600
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSDataGridLib.DataGrid dgKonversi2 
            Height          =   4695
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   8281
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
                  LCID            =   1033
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
                  LCID            =   1033
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Shift"
            Height          =   210
            Left            =   3720
            TabIndex        =   40
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Pegawai"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.TextBox txtJamPulangAkhir 
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
         TabIndex        =   6
         Top             =   3660
         Width           =   1575
      End
      Begin VB.TextBox txtJamPulangAwal 
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
         TabIndex        =   5
         Top             =   3180
         Width           =   1575
      End
      Begin VB.TextBox txtJamMasukAkhir 
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
         TabIndex        =   4
         Top             =   2700
         Width           =   1575
      End
      Begin VB.TextBox txtJamMasukAwal 
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
         TabIndex        =   3
         Top             =   2220
         Width           =   1575
      End
      Begin VB.TextBox txtJamPulang 
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
         TabIndex        =   2
         Top             =   1740
         Width           =   1575
      End
      Begin VB.TextBox txtJamMasuk 
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
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   -74400
         TabIndex        =   19
         Top             =   660
         Width           =   7575
         Begin VB.TextBox txtKdRuangan 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   720
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Frame fraRuangan 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Visible         =   0   'False
            Width           =   3495
            Begin MSDataGridLib.DataGrid dgRuangan 
               Height          =   3495
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   6165
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
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
         End
         Begin VB.TextBox txtRuangan 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   3375
         End
         Begin MSDataListLib.DataCombo dcShift 
            Height          =   315
            Left            =   3720
            TabIndex        =   12
            Top             =   600
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSDataGridLib.DataGrid dgKonversi 
            Height          =   3375
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   5953
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
                  LCID            =   1033
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
                  LCID            =   1033
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
         Begin MSDataListLib.DataCombo dcShift22 
            Height          =   315
            Left            =   3720
            TabIndex        =   41
            Top             =   1200
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcShift23 
            Height          =   315
            Left            =   3720
            TabIndex        =   43
            Top             =   1800
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Shift 3"
            Height          =   195
            Left            =   3720
            TabIndex        =   44
            Top             =   1560
            Width           =   450
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Shift 2"
            Height          =   195
            Left            =   3720
            TabIndex        =   42
            Top             =   960
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Ruangan"
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Shift 1"
            Height          =   195
            Left            =   3720
            TabIndex        =   22
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.TextBox txtKdshift 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   16
         Top             =   4140
         Visible         =   0   'False
         Width           =   855
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
         Height          =   315
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   0
         Top             =   780
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dgshift 
         Height          =   5775
         Left            =   3360
         TabIndex        =   14
         Top             =   780
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   10186
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jam Pulang Akhir"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   3660
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jam Pulang Awal"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   3180
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Jam Masuk Akhir"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jam Masuk Awal"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   2220
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jam Pulang"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1740
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jam Masuk"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1260
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   195
         Left            =   840
         TabIndex        =   18
         Top             =   4140
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Shift"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   780
         Width           =   780
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   24
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
      Left            =   7440
      Picture         =   "frmMasterShift.frx":0D1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterShift.frx":1AA6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterShift.frx":4467
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMasterShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub subDcSource()
    strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja order by NamaShift"
    Call msubDcSource(dcShift, rs, strSQL)

    strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja order by NamaShift"
    Call msubDcSource(dcShift2, rs, strSQL)

    strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja order by NamaShift"
    Call msubDcSource(dcShift22, rs, strSQL)

    strSQL = "SELECT KdShift, NamaShift FROM ShiftKerja order by NamaShift"
    Call msubDcSource(dcShift23, rs, strSQL)
End Sub

Sub sp_simpan()
    Select Case sstDataPenunjang.Tab
        Case 0 'shift
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdShift", adVarChar, adParamInput, 2, Trim(txtKdshift.Text))
                .Parameters.Append .CreateParameter("NamaShift", adVarChar, adParamInput, 20, Trim(txtnamashift.Text))
                .Parameters.Append .CreateParameter("JamMasuk", adChar, adParamInput, 5, Trim(txtJamMasuk.Text))
                .Parameters.Append .CreateParameter("JamPulang", adChar, adParamInput, 5, Trim(txtJamPulang.Text))
                .Parameters.Append .CreateParameter("JamMasukAwal", adChar, adParamInput, 5, Trim(txtJamMasukAwal.Text))
                .Parameters.Append .CreateParameter("JamMasukAkhir", adChar, adParamInput, 5, Trim(txtJamMasukAkhir.Text))
                .Parameters.Append .CreateParameter("JamPulangAwal", adChar, adParamInput, 5, Trim(txtJamPulangAwal.Text))
                .Parameters.Append .CreateParameter("JamPulangAkhir", adChar, adParamInput, 5, Trim(txtJamPulangAkhir.Text))
                .Parameters.Append .CreateParameter("OutputKdShift", adVarChar, adParamOutput, 2, Null)
                .Parameters.Append .CreateParameter("status", adVarChar, adParamInput, 1, "A")

                .ActiveConnection = dbConn
                .CommandText = "AUD_ShiftKerja"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data", vbExclamation, "Validasi"
                Else
                    If Not IsNull(.Parameters("OutputKdShift").Value) Then txtKdshift = .Parameters("OutputKdShift").Value
                End If
                Call deleteADOCommandParameters(dbcmd)
            End With
            cmdBatal_Click

        Case 1 'konversi shift to ruangan
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdShift", adVarChar, adParamInput, 2, dcShift.BoundText)
                .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, txtKdRuangan.Text)
                .Parameters.Append .CreateParameter("KdShift2", adVarChar, adParamInput, 2, IIf(dcShift22.BoundText = "", Null, dcShift22.BoundText))
                .Parameters.Append .CreateParameter("KdShift3", adVarChar, adParamInput, 2, IIf(dcShift23.BoundText = "", Null, dcShift23.BoundText))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")

                .ActiveConnection = dbConn
                .CommandText = "AUD_ConvertShiftToRuangan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)
            End With
            cmdBatal_Click

        Case 2 'konversi shift to pegawai
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIDPegawai.Text)
                .Parameters.Append .CreateParameter("KdShift", adVarChar, adParamInput, 2, dcShift2.BoundText)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")

                .ActiveConnection = dbConn
                .CommandText = "AUD_ConvertShiftToPegawai"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)
            End With
            cmdBatal_Click

    End Select
End Sub

Private Sub cmdBatal_Click()
    Select Case sstDataPenunjang.Tab
        Case 0
            Call subKosong

        Case 1
            Call subKosong

        Case 2
            Call subKosong
    End Select
End Sub

Private Sub cmdHapus_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 'shift
            Set rs = Nothing
            strSQL = "delete shiftkerja where kdshift = '" & txtKdshift.Text & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 1 'konversi shift
            Set rs = Nothing
            strSQL = "delete convertshifttoruangan where Kdshift = '" & dcShift.BoundText & "' and kdruangan= '" & txtKdRuangan.Text & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 2 'konversi pegawai
            Set rs = Nothing
            strSQL = "delete convertshifttopegawai where Kdshift = '" & dcShift2.BoundText & "' and IdPegawai= '" & txtIDPegawai.Text & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

    End Select
    Call subLoadGridSource
    Call subKosong
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan
    Select Case sstDataPenunjang.Tab
        Case 0 'shift
            If Periksa("text", txtnamashift, "Nama Shift kosong") = False Then Exit Sub
            If Periksa("text", txtJamMasuk, "Jam Masuk kosong") = False Then Exit Sub
            If Periksa("text", txtJamPulang, "Jam Pulang kosong") = False Then Exit Sub
            If Periksa("text", txtJamMasukAwal, "Toleransi Jam Masuk Awal kosong") = False Then Exit Sub
            If Periksa("text", txtJamMasukAkhir, "Toleransi Jam Masuk Akhir kosong") = False Then Exit Sub
            If Periksa("text", txtJamPulangAwal, "Toleransi Jam Pulang Awal kosong") = False Then Exit Sub
            If Periksa("text", txtJamPulangAkhir, "Toleransi Jam Pulang Akhir kosong") = False Then Exit Sub
            Call sp_simpan

        Case 1  ' konversi shift to ruangan
            If Periksa("text", txtRuangan, "Nama Ruangan kosong") = False Then Exit Sub
            If Periksa("datacombo", dcShift, "Silahkan isi Nama Shift") = False Then Exit Sub
            Call sp_simpan

        Case 2  ' konversi shift to pegawai
            If Periksa("text", txtPegawai, "Pilih nama pegawai") = False Then Exit Sub
            If Periksa("datacombo", dcShift2, "Silahkan isi Nama Shift") = False Then Exit Sub
            Call sp_simpan

    End Select
    Call subLoadGridSource
    Exit Sub
errSimpan:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcShift_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcShift22.SetFocus
End Sub

Private Sub dcShift2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dcShift22_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcShift23.SetFocus
End Sub

Private Sub dcShift23_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dgKonversi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgKonversi
        txtRuangan.Text = .Columns(0).Value
        dcShift.BoundText = .Columns(3).Value
        If .Columns(4).Value = Null Then
            dcShift22.BoundText = ""
        Else
            dcShift22.BoundText = .Columns(4).Value
        End If
        If .Columns(5).Value = Null Then
            dcShift23.BoundText = ""
        Else
            dcShift23.BoundText = .Columns(5).Value
        End If

        txtKdRuangan.Text = .Columns(2).Value
    End With
    fraRuangan.Visible = False
End Sub

Private Sub dgKonversi2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgKonversi2
        txtPegawai.Text = .Columns(0).Value
        txtIDPegawai.Text = .Columns(2).Value
        dcShift2.BoundText = .Columns(3).Value
    End With
    fraPegawai.Visible = False
End Sub

Private Sub dgRuangan_DblClick()
    Call dgRuangan_KeyPress(13)
End Sub

Private Sub dgPegawai_DblClick()
    Call dgPegawai_KeyPress(13)
End Sub

Private Sub dgshift_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdshift.Text = dgshift.Columns(0).Value
    txtnamashift.Text = dgshift.Columns(1).Value
    txtJamMasuk = dgshift.Columns(2).Value
    txtJamPulang.Text = dgshift.Columns(3).Value
    txtJamMasukAwal.Text = dgshift.Columns(4).Value
    txtJamMasukAkhir.Text = dgshift.Columns(5).Value
    txtJamPulangAwal.Text = dgshift.Columns(6).Value
    txtJamPulangAkhir.Text = dgshift.Columns(7).Value

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
        Case 0 'shift
            txtKdshift.Text = ""
            txtnamashift.Text = ""
            txtJamMasuk.Text = ""
            txtJamPulang.Text = ""
            txtJamMasukAwal.Text = ""
            txtJamMasukAkhir.Text = ""
            txtJamPulangAwal.Text = ""
            txtJamPulangAkhir.Text = ""

        Case 1 'konversi shift
            txtRuangan.Text = ""
            dcShift.Text = ""
            dcShift22.Text = ""
            dcShift23.Text = ""
            fraRuangan.Visible = False

        Case 2 'konversi pegawai
            txtIDPegawai.Text = ""
            txtPegawai.Text = ""
            dcShift2.BoundText = ""
            fraPegawai.Visible = False
    End Select
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call subDcSource
    Call subLoadGridSource
    Call cmdBatal_Click
End Sub

Sub subLoadGridSource()
    Select Case sstDataPenunjang.Tab
        Case 0 ' shift
            Set rs = Nothing
            strSQL = "select * from ShiftKerja order by namashift"
            Call msubRecFO(rs, strSQL)
            Set dgshift.DataSource = rs
            With dgshift
                .Columns(0).Width = 0
                .Columns(1).Width = 1500
                .Columns(2).Width = 1500
                .Columns(3).Width = 1500
                .Columns(4).Width = 0
                .Columns(5).Width = 0
                .Columns(6).Width = 0
                .Columns(7).Width = 0
            End With

        Case 1  'konversi shift to ruangan
            Set rs = Nothing
            strSQL = "select * from V_KonversiShiftToRuangan  order by namaruangan"
            Call msubRecFO(rs, strSQL)
            Set dgKonversi.DataSource = rs
            With dgKonversi
                .Columns(0).Width = 2500
                .Columns(1).Width = 1500
                .Columns(2).Width = 0
                .Columns(3).Width = 0
                .Columns(4).Width = 0
                .Columns(5).Width = 0
                .Columns(6).Width = 1500
                .Columns(7).Width = 1500
            End With

        Case 2  'konversi shift to pegawai
            Set rs = Nothing
            strSQL = "select * from V_KonversiShiftToPegawai  order by namalengkap"
            Call msubRecFO(rs, strSQL)
            Set dgKonversi2.DataSource = rs
            With dgKonversi2
                .Columns(0).Width = 5000
                .Columns(1).Width = 1500
                .Columns(2).Width = 0
                .Columns(3).Width = 0
            End With
    End Select
End Sub

Private Sub txtJamMasuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJamPulang.SetFocus
End Sub

Private Sub txtJamMasukAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJamPulangAwal.SetFocus
End Sub

Private Sub txtJamMasukAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJamMasukAkhir.SetFocus
End Sub

Private Sub txtJamPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJamMasukAwal.SetFocus
End Sub

Private Sub txtJamPulangAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtJamPulangAwal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJamPulangAkhir.SetFocus
End Sub

Private Sub txtnamashift_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJamMasuk.SetFocus
End Sub

Private Sub txtRuangan_Change()
    On Error GoTo errLoad
    strFilterPasien = "WHERE NamaRuangan like '%" & txtRuangan.Text & "%' order by NamaRuangan"
    fraRuangan.Visible = True
    Call subLoadRuangan
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub txtPegawai_Change()
    On Error GoTo errLoad
    strFilterPasien = "WHERE NamaLengkap like '%" & txtPegawai.Text & "%' order by NamaLengkap"
    fraPegawai.Visible = True
    Call subLoadPegawai
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub subLoadRuangan()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select KdRuangan, NamaRuangan from Ruangan " & strFilterPasien
    Call msubRecFO(rs, strSQL)
    Set dgRuangan.DataSource = rs
    With dgRuangan
        .Columns(0).Width = 0
        .Columns(1).Width = 2500
    End With
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub subLoadPegawai()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select IdPegawai, NamaLengkap from DataPEgawai " & strFilterPasien
    Call msubRecFO(rs, strSQL)
    Set dgPegawai.DataSource = rs
    With dgPegawai
        .Columns(0).Width = 0
        .Columns(1).Width = 2500
    End With
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub dgRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRuangan.Text = dgRuangan.Columns(1).Value
        txtKdRuangan.Text = dgRuangan.Columns(0).Value
        fraRuangan.Visible = False
        dcShift.SetFocus
    End If
    If KeyAscii = 27 Then
        fraRuangan.Visible = False
    End If
End Sub

Private Sub dgPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPegawai.Text = dgPegawai.Columns(1).Value
        txtIDPegawai.Text = dgPegawai.Columns(0).Value
        fraPegawai.Visible = False
        dcShift2.SetFocus
    End If
    If KeyAscii = 27 Then
        fraPegawai.Visible = False
    End If
End Sub

Private Sub txtRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcShift.SetFocus
End Sub

Private Sub txtPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcShift2.SetFocus
End Sub
