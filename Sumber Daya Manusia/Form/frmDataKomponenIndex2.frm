VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmDataKomponenIndex2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Master Komponen Index"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "frmDataKomponenIndex2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   10575
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
      Left            =   4800
      TabIndex        =   34
      Top             =   8175
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
      Left            =   9120
      TabIndex        =   33
      Top             =   8175
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
      Left            =   6240
      TabIndex        =   32
      Top             =   8175
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
      Left            =   1440
      TabIndex        =   31
      Top             =   8160
      Visible         =   0   'False
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
      Left            =   7680
      TabIndex        =   30
      Top             =   8175
      Width           =   1335
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   6975
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   12303
      _Version        =   393216
      TabHeight       =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Jenis Komponen Index"
      TabPicture(0)   =   "frmDataKomponenIndex2.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameJenisKomp"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Komponen Index"
      TabPicture(1)   =   "frmDataKomponenIndex2.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameKomponen"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Detail Komponen Index"
      TabPicture(2)   =   "frmDataKomponenIndex2.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameDetailKomponen"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame frameJenisKomp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5925
         Left            =   240
         TabIndex        =   25
         Top             =   750
         Width           =   9615
         Begin VB.TextBox txtCariJK 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   600
            MaxLength       =   50
            TabIndex        =   35
            Top             =   5400
            Width           =   4695
         End
         Begin VB.TextBox txtKodeJenisKomp 
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
            Height          =   330
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   0
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtJenisKomp 
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
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   1
            Top             =   840
            Width           =   6015
         End
         Begin MSDataGridLib.DataGrid dgJenisKomp 
            Height          =   3855
            Left            =   240
            TabIndex        =   26
            Top             =   1320
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cari "
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
            TabIndex        =   36
            Top             =   5400
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Kode "
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
            Left            =   840
            TabIndex        =   28
            Top             =   420
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Komponen Index"
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
            Left            =   840
            TabIndex        =   27
            Top             =   900
            Width           =   1860
         End
      End
      Begin VB.Frame frameKomponen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5925
         Left            =   -74760
         TabIndex        =   19
         Top             =   750
         Width           =   9855
         Begin VB.TextBox txtCariKK 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   600
            MaxLength       =   50
            TabIndex        =   37
            Top             =   5400
            Width           =   4695
         End
         Begin VB.TextBox txtNamaKomponen 
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
            Left            =   2640
            MaxLength       =   50
            TabIndex        =   8
            Top             =   720
            Width           =   6495
         End
         Begin VB.TextBox txtKodeKomp 
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
            Height          =   330
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dcJenisKomp 
            Height          =   315
            Left            =   2640
            TabIndex        =   10
            Top             =   1200
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataGridLib.DataGrid dgKomponenIndex 
            Height          =   3495
            Left            =   240
            TabIndex        =   20
            Top             =   1800
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   6165
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgJenisKomp2 
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   1800
            Visible         =   0   'False
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   661
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cari "
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
            TabIndex        =   38
            Top             =   5400
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Kode "
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
            Left            =   600
            TabIndex        =   24
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Komponen Index"
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
            Left            =   600
            TabIndex        =   23
            Top             =   780
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Komponen Index"
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
            Left            =   600
            TabIndex        =   22
            Top             =   1260
            Width           =   1860
         End
      End
      Begin VB.Frame frameDetailKomponen 
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
         Left            =   -74760
         TabIndex        =   11
         Top             =   750
         Width           =   9855
         Begin VB.TextBox txtCariDK 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   600
            MaxLength       =   50
            TabIndex        =   39
            Top             =   5400
            Width           =   4695
         End
         Begin VB.TextBox txtRateIndex 
            Alignment       =   1  'Right Justify
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
            Left            =   5400
            TabIndex        =   6
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtNilaiIndexStandar 
            Alignment       =   1  'Right Justify
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
            Left            =   2760
            TabIndex        =   5
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtKodeDetailKomp 
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
            Height          =   330
            Left            =   2760
            MaxLength       =   6
            TabIndex        =   3
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtNamaDetailKomp 
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
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1080
            Width           =   6495
         End
         Begin MSDataListLib.DataCombo dcKomponenIndex 
            Height          =   315
            Left            =   2760
            TabIndex        =   2
            Top             =   360
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataGridLib.DataGrid dgDetailKomponenIndex 
            Height          =   2295
            Left            =   240
            TabIndex        =   12
            Top             =   3000
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgKomponenIndex2 
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   3600
            Visible         =   0   'False
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   450
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcJabatan 
            Height          =   315
            Left            =   2760
            TabIndex        =   41
            Top             =   1800
            Width           =   3735
            _ExtentX        =   6588
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
         Begin MSDataListLib.DataCombo dcPendidikan 
            Height          =   315
            Left            =   2760
            TabIndex        =   43
            Top             =   2160
            Width           =   3735
            _ExtentX        =   6588
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
         Begin MSDataListLib.DataCombo dcInstalasi 
            Height          =   315
            Left            =   2760
            TabIndex        =   45
            Top             =   2520
            Width           =   3735
            _ExtentX        =   6588
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
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Nama Instalasi"
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
            Left            =   600
            TabIndex        =   46
            Top             =   2580
            Width           =   1140
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Nama Pendidikan"
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
            Left            =   600
            TabIndex        =   44
            Top             =   2220
            Width           =   1380
         End
         Begin VB.Label Label22 
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
            Left            =   600
            TabIndex        =   42
            Top             =   1860
            Width           =   1140
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cari "
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
            TabIndex        =   40
            Top             =   5400
            Width           =   345
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Rate Index"
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
            Left            =   4320
            TabIndex        =   18
            Top             =   1500
            Width           =   900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Index Standar"
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
            Left            =   600
            TabIndex        =   17
            Top             =   1440
            Width           =   1515
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Komponen Index"
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
            Left            =   600
            TabIndex        =   16
            Top             =   420
            Width           =   1410
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Detail Komponen Index"
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
            Left            =   600
            TabIndex        =   15
            Top             =   1140
            Width           =   1920
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Kode "
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
            Left            =   600
            TabIndex        =   14
            Top             =   780
            Width           =   480
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   29
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataKomponenIndex2.frx":0D1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8760
      Picture         =   "frmDataKomponenIndex2.frx":36DF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataKomponenIndex2.frx":4467
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDataKomponenIndex2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilter As String

Private Sub cmdSimpan_Click()
    Call sp_simpan("A")
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsimpan.SetFocus
End Sub

Private Sub dcJenisKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsimpan.SetFocus
End Sub

Private Sub dcKomponenIndex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaDetailKomp.SetFocus
End Sub

Private Sub dgDetailKomponenIndex_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDetailKomponenIndex
    WheelHook.WheelHook dgDetailKomponenIndex
End Sub

Private Sub dgJenisKomp_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgJenisKomp
    WheelHook.WheelHook dgJenisKomp
End Sub

Private Sub dgJenisKomp2_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgJenisKomp2
    WheelHook.WheelHook dgJenisKomp2
End Sub

Private Sub dgKomponenIndex_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKomponenIndex
    WheelHook.WheelHook dgKomponenIndex
End Sub

Private Sub dgKomponenIndex2_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKomponenIndex2
    WheelHook.WheelHook dgKomponenIndex2
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call SetComboJenisKomp
    Call SetComboNamaKomponenIndex
    Call SetComboJabatan
    Call SetComboPendidikan
    Call SetComboInstalasi
    Call tampilData

    sstDataPenunjang.Tab = 0
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    Call clear
    Call tampilData
    Call SetComboJenisKomp
    Call SetComboNamaKomponenIndex
    Call SetComboJabatan
    Call SetComboPendidikan
    Call SetComboInstalasi
End Sub

Private Sub sp_simpan(f_Status As String)
    On Error GoTo errSimpan
    Select Case sstDataPenunjang.Tab
        Case 0

            If Periksa("text", txtJenisKomp, "Jenis Komponen Index kosong") = False Then Exit Sub
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdJenisKomponenIndex", adChar, adParamInput, 2, Trim(txtKodeJenisKomp))
                .Parameters.Append .CreateParameter("JenisKomponenIndex", adVarChar, adParamInput, 50, Trim(txtJenisKomp))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_JenisKomponenIndex"
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

        Case 1

            If Periksa("text", txtNamaKomponen, "Nama Komponen Index kosong") = False Then Exit Sub
            If Periksa("datacombo", dcJenisKomp, "Silahkan pilih Jenis Komponen Index") = False Then Exit Sub
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdKomponenIndex", adVarChar, adParamInput, 4, Trim(txtKodeKomp))
                .Parameters.Append .CreateParameter("KomponenIndex", adVarChar, adParamInput, 50, Trim(txtNamaKomponen))
                .Parameters.Append .CreateParameter("KdJenisKomponenIndex", adChar, adParamInput, 2, dcJenisKomp.BoundText)
                .Parameters.Append .CreateParameter("OutKode", adVarChar, adParamInputOutput, 4, Null)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_KomponenIndex"
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

        Case 2

            If Periksa("datacombo", dcKomponenIndex, "Silahkan pilih Komponen Index") = False Then Exit Sub
            If Periksa("text", txtNamaDetailKomp, "Detail Komponen Index kosong") = False Then Exit Sub
            If Periksa("text", txtNilaiIndexStandar, "Nilai Index Standar kosong") = False Then Exit Sub
            If Periksa("text", txtRateIndex, "Rate Index kosong") = False Then Exit Sub

            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdDetailKomponenIndex", adVarChar, adParamInput, 6, Trim(txtKodeDetailKomp))
                .Parameters.Append .CreateParameter("KdKomponenIndex", adVarChar, adParamInput, 4, dcKomponenIndex.BoundText)
                .Parameters.Append .CreateParameter("DetailKomponenIndex", adVarChar, adParamInput, 50, Trim(txtNamaDetailKomp))
                .Parameters.Append .CreateParameter("NilaiIndexStandar", adInteger, adParamInput, , Val(txtNilaiIndexStandar))
                .Parameters.Append .CreateParameter("RateIndex", adInteger, adParamInput, , Val(txtRateIndex))
                .Parameters.Append .CreateParameter("KdJabatan", adVarChar, adParamInput, 5, IIf(dcJabatan.Text = "", Null, Trim(dcJabatan.BoundText)))
                .Parameters.Append .CreateParameter("KdPendidikan", adChar, adParamInput, 4, IIf(dcPendidikan.Text = "", Null, dcPendidikan.BoundText))
                .Parameters.Append .CreateParameter("KdInstalasi", adChar, adParamInput, 2, IIf(dcInstalasi.Text = "", Null, dcInstalasi.BoundText))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_DetailKomponenIndex"
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

    End Select
    clear
    tampilData
    Exit Sub
errSimpan:
    msubPesanError
End Sub

Private Sub msg()
    MsgBox "Data Berhasil Dihapus", vbInformation, "Informasi"
End Sub

Private Sub cmdHapus_Click()
    Select Case sstDataPenunjang.Tab
        Case 0
            If txtKodeJenisKomp.Text = "" Then
                MsgBox "Ma'af Data Belum Dipilih!", vbInformation + vbOKOnly, "Pesan Data Kosong"
            ElseIf txtKodeJenisKomp.Text <> "" Then
                Set rs = Nothing
                strSQL = "select * from KomponenIndex where KdJenisKomponenIndex = '" & txtKodeJenisKomp & "'"
                rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                If rs.RecordCount <> 0 Then
                    MsgBox "Data tidak dapat dihapus"
                    dgJenisKomp.SetFocus
                    Exit Sub
                End If

                Set rs = Nothing
                strSQL = "delete JenisKomponenIndex  where KdJenisKomponenIndex = '" & txtKodeJenisKomp & "'"
                rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
                Set rs = Nothing
                Call msg
            End If
        Case 1
            If txtKodeKomp.Text = "" Then
                MsgBox "Ma'af Data Belum Dipilih!", vbInformation + vbOKOnly, "Pesan Data Kosong"
            ElseIf txtKodeKomp.Text <> "" Then
                Set rs = Nothing
                strSQL = "select * from DetailKomponenIndex where KdKomponenIndex = '" & txtKodeKomp & "'"
                rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                If rs.RecordCount <> 0 Then
                    MsgBox " data tidak dapat dihapus"
                    dgKomponenIndex.SetFocus
                    Exit Sub
                End If

                Set rs = Nothing
                strSQL = "delete KomponenIndex where KdKomponenIndex= '" & txtKodeKomp & "'"
                rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
                Set rs = Nothing
                Call msg
            End If
        Case 2
            If dcKomponenIndex.Text = "" Then
                MsgBox "Ma'af Data Belum Dipilih!", vbInformation + vbOKOnly, "Pesan Data Kosong"
            ElseIf dcKomponenIndex.Text <> "" Then
                Set rs = Nothing
                strSQL = "select * from TotalScoreIndex where KdDetailKomponenIndex = '" & txtKodeDetailKomp & "'"
                rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                If rs.RecordCount <> 0 Then
                    MsgBox "data tidak dapat Di Hapus"
                    dgDetailKomponenIndex.SetFocus
                    Exit Sub
                End If

                Set rs = Nothing
                strSQL = "delete DetailKomponenIndex where KdDetailKomponenIndex = '" & txtKodeDetailKomp & "'"
                rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
                Set rs = Nothing
                Call msg
            End If

    End Select
    clear
    tampilData
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Sub clear()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0
            txtKodeJenisKomp.Text = ""
            txtJenisKomp.Text = ""
            txtJenisKomp.SetFocus

        Case 1
            txtKodeKomp.Text = ""
            txtNamaKomponen.Text = ""
            dcJenisKomp.Text = ""
            txtNamaKomponen.SetFocus

        Case 2
            txtKodeDetailKomp.Text = ""
            txtNamaDetailKomp.Text = ""
            txtNilaiIndexStandar.Text = ""
            txtRateIndex.Text = ""
            dcKomponenIndex.Text = ""
            dcJabatan.BoundText = ""
            dcPendidikan.BoundText = ""
            dcInstalasi.BoundText = ""
            dcKomponenIndex.SetFocus

    End Select
End Sub

Sub tampilData()
    On Error GoTo errTampil
    Select Case sstDataPenunjang.Tab
        Case 0

            Set rs = Nothing
            strSQL = "select * from JenisKomponenIndex where JenisKomponenIndex like '%" & txtCariJK.Text & "%' order by KdJenisKomponenIndex"
            Call msubRecFO(rs, strSQL)
            Set dgJenisKomp.DataSource = rs
            With dgJenisKomp
                .Columns(0).Caption = "Kode Jenis"
                .Columns(1).Caption = "Jenis Komponen Index"
                .Columns(0).Width = 1500
                .Columns(1).Width = 6000
            End With

        Case 1
            Set rs = Nothing
            strSQL = "select * from V_DataKomponenIndex where KomponenIndex like '%" & txtCariKK.Text & "%' "
            Call msubRecFO(rs, strSQL)
            Set dgKomponenIndex.DataSource = rs
            With dgKomponenIndex
                .Columns(0).Caption = "Kode"
                .Columns(1).Caption = "Komponen Index"
                .Columns(2).Caption = "Jenis Komponen Index"
                .Columns(0).Width = 1000
                .Columns(1).Width = 4500
                .Columns(2).Width = 3000
                .Columns(3).Width = 0
            End With

        Case 2
            Set rs = Nothing
            strSQL = "select * from V_DetailKomponenIndex where DetailKomponenIndex like '%" & txtCariDK.Text & "%' "
            Call msubRecFO(rs, strSQL)
            Set dgDetailKomponenIndex.DataSource = rs
            With dgDetailKomponenIndex
                .Columns(0).Caption = "KD KOMPONEN INDEX"
                .Columns(1).Caption = "Komponen Index"
                .Columns(2).Caption = "Kode"
                .Columns(3).Caption = "Detail Komponen Index"
                .Columns(4).Caption = "Nilai Index Standar"
                .Columns(5).Caption = "Rate Index"
                .Columns(6).Caption = "Nama Jabatan"
                .Columns(7).Caption = "Pendidikan"
                .Columns(8).Caption = "Instalasi"
                .Columns(0).Width = 0
                .Columns(1).Width = 1500
                .Columns(2).Width = 0
                .Columns(3).Width = 1750
                .Columns(4).Width = 2000
                .Columns(5).Width = 1200
                .Columns(6).Width = 2000
                .Columns(7).Width = 2000
                .Columns(8).Width = 2000
                .Columns(9).Width = 0
                .Columns(10).Width = 0
                .Columns(11).Width = 0
            End With

    End Select
    Exit Sub
errTampil:
    Call msubPesanError
End Sub

Private Sub dgJenisKomp_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKodeJenisKomp.Text = dgJenisKomp.Columns(0).Value
    txtJenisKomp.Text = dgJenisKomp.Columns(1).Value
End Sub

Private Sub dgKomponenIndex_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    dcJenisKomp.BoundText = dgKomponenIndex.Columns(3).Value
    txtKodeKomp.Text = dgKomponenIndex.Columns(0).Value
    txtNamaKomponen.Text = dgKomponenIndex.Columns(1).Value
End Sub

Private Sub dgDetailKomponenIndex_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    dcKomponenIndex.BoundText = dgDetailKomponenIndex.Columns(0).Value
    txtKodeDetailKomp.Text = dgDetailKomponenIndex.Columns(2).Value
    txtNamaDetailKomp.Text = dgDetailKomponenIndex.Columns(3).Value
    txtNilaiIndexStandar.Text = dgDetailKomponenIndex.Columns(4).Value
    txtRateIndex.Text = dgDetailKomponenIndex.Columns(5).Value
    If IsNull(dgDetailKomponenIndex.Columns(9).Value) Then dcJabatan.BoundText = "" Else dcJabatan.BoundText = dgDetailKomponenIndex.Columns(9).Value
    If IsNull(dgDetailKomponenIndex.Columns(10).Value) Then dcPendidikan.BoundText = "" Else dcPendidikan.BoundText = dgDetailKomponenIndex.Columns(10).Value
    If IsNull(dgDetailKomponenIndex.Columns(11).Value) Then dcInstalasi.BoundText = "" Else dcInstalasi.BoundText = dgDetailKomponenIndex.Columns(11).Value
End Sub

Sub SetComboJenisKomp()
    strSQL = "SELECT * FROM JenisKomponenIndex order by JenisKomponenIndex"
    Call msubDcSource(dcJenisKomp, rs, strSQL)
End Sub

Sub SetComboNamaKomponenIndex()
    strSQL = "Select KomponenIndex.KdKomponenIndex, KomponenIndex.KomponenIndex, JenisKomponenIndex.JenisKomponenIndex" & _
    " FROM KomponenIndex INNER JOIN JenisKomponenIndex ON KomponenIndex.KdJenisKomponenIndex = JenisKomponenIndex.KdJenisKomponenIndex" & _
    " WHERE KomponenIndex like '%" & dcKomponenIndex.Text & "%'"
    Call msubDcSource(dcKomponenIndex, rs, strSQL)
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call tampilData
End Sub

Private Sub txtCariDK_Change()
    Call tampilData
End Sub

Private Sub txtCariJK_Change()
    Call tampilData
End Sub

Private Sub txtCariKK_Change()
    Call tampilData
End Sub

Private Sub txtJenisKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsimpan.SetFocus
End Sub

Private Sub txtKodeJenisKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtJenisKomp.SetFocus
End Sub

Private Sub txtKodeKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNamaKomponen.SetFocus
End Sub

Private Sub txtKodeDetailKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNamaDetailKomp.SetFocus
End Sub

Private Sub txtNamaDetailKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNilaiIndexStandar.SetFocus
End Sub

Private Sub txtNamaKomponen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisKomp.SetFocus
End Sub

Private Sub txtNilaiIndexStandar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtRateIndex.SetFocus
End Sub

Private Sub dcJabatan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPendidikan.SetFocus
End Sub

Private Sub dcPendidikan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcInstalasi.SetFocus
End Sub

Sub SetComboJabatan()
    Set rs = Nothing
    strSQL = "Select KdJabatan, NamaJabatan from Jabatan order by NamaJabatan"
    Call msubDcSource(dcJabatan, rs, strSQL)

End Sub

Sub SetComboPendidikan()
    Set rs = Nothing
    strSQL = "Select KdPendidikan, Pendidikan from Pendidikan order by Pendidikan"
    Call msubDcSource(dcPendidikan, rs, strSQL)
End Sub

Sub SetComboInstalasi()
    Set rs = Nothing
    strSQL = "Select KdInstalasi, NamaInstalasi from Instalasi order by NamaInstalasi "
    Call msubDcSource(dcInstalasi, rs, strSQL)

End Sub

Private Sub txtRateIndex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJabatan.SetFocus
End Sub
