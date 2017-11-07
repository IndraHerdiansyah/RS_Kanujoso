VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmDataKomponenIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Komponen Index"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "frmDataKomponenIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7770
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
      Left            =   3240
      TabIndex        =   59
      Top             =   7335
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
      Left            =   9000
      TabIndex        =   58
      Top             =   7335
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
      Left            =   4680
      TabIndex        =   57
      Top             =   7320
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
      Left            =   7560
      TabIndex        =   56
      Top             =   7335
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
      Left            =   6120
      TabIndex        =   55
      Top             =   7335
      Width           =   1335
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   6135
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   5
      TabHeight       =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Jenis Komponen Index"
      TabPicture(0)   =   "frmDataKomponenIndex.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameJenisKomp"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Komponen Index"
      TabPicture(1)   =   "frmDataKomponenIndex.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameKomponen"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Detail Komponen Index"
      TabPicture(2)   =   "frmDataKomponenIndex.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameDetailKomponen"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Konversi Jabatan"
      TabPicture(3)   =   "frmDataKomponenIndex.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Konversi Pendidikan"
      TabPicture(4)   =   "frmDataKomponenIndex.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).ControlCount=   1
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
         Height          =   4695
         Left            =   -74760
         TabIndex        =   42
         Top             =   1200
         Width           =   9735
         Begin VB.TextBox txtRIPendidikan 
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
            Left            =   5280
            TabIndex        =   45
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtNISPendidikan 
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
            Left            =   2520
            TabIndex        =   44
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   5520
            TabIndex        =   43
            Top             =   4800
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dcPendidikan 
            Height          =   315
            Left            =   2520
            TabIndex        =   46
            Top             =   360
            Width           =   5055
            _ExtentX        =   8916
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
         Begin MSDataGridLib.DataGrid dgPendidikan 
            Height          =   2535
            Left            =   240
            TabIndex        =   47
            Top             =   1920
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4471
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
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
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   2280
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
         Begin MSDataListLib.DataCombo dcDetailKomponenIndexPendidikan 
            Height          =   315
            Left            =   2520
            TabIndex        =   49
            Top             =   840
            Width           =   5055
            _ExtentX        =   8916
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
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Rate Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4200
            TabIndex        =   53
            Top             =   1440
            Width           =   945
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Index Standar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   1440
            Width           =   1605
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nama Pendidikan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   420
            Width           =   1455
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Detail Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   960
            Width           =   1980
         End
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
         Height          =   4695
         Left            =   -74760
         TabIndex        =   35
         Top             =   1200
         Width           =   9735
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   5520
            TabIndex        =   36
            Top             =   4680
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtNISJabatan 
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
            Left            =   2520
            TabIndex        =   3
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtRIJabatan 
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
            TabIndex        =   4
            Top             =   1320
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dcJabatan 
            Height          =   315
            Left            =   2520
            TabIndex        =   1
            Top             =   360
            Width           =   5055
            _ExtentX        =   8916
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
         Begin MSDataGridLib.DataGrid dgJabatan 
            Height          =   2655
            Left            =   240
            TabIndex        =   37
            Top             =   1800
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4683
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
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
         Begin MSDataListLib.DataCombo dcDetailKomponenIndexJabatan 
            Height          =   315
            Left            =   2520
            TabIndex        =   2
            Top             =   840
            Width           =   5055
            _ExtentX        =   8916
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Detail Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   960
            Width           =   1980
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Nama Jabatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Index Standar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   1440
            Width           =   1605
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Rate Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3960
            TabIndex        =   38
            Top             =   1440
            Width           =   945
         End
      End
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
         Height          =   4575
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   9615
         Begin VB.TextBox txtKodeJenisKomp 
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
            MaxLength       =   2
            TabIndex        =   12
            Top             =   360
            Width           =   6495
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
            TabIndex        =   13
            Top             =   840
            Width           =   6495
         End
         Begin MSDataGridLib.DataGrid dgJenisKomp 
            Height          =   3015
            Left            =   240
            TabIndex        =   32
            Top             =   1320
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   5318
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Kode Jenis Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   420
            Width           =   2385
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   900
            Width           =   1920
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
         Height          =   4695
         Left            =   -74760
         TabIndex        =   24
         Top             =   1200
         Width           =   9615
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
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1320
            Width           =   6975
         End
         Begin VB.TextBox txtKodeKomp 
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
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   10
            Top             =   840
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   7680
            TabIndex        =   25
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dcJenisKomp 
            Height          =   315
            Left            =   2400
            TabIndex        =   9
            Top             =   360
            Width           =   3135
            _ExtentX        =   5530
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
            Height          =   2655
            Left            =   240
            TabIndex        =   26
            Top             =   1800
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4683
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
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
            TabIndex        =   27
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Kode Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   900
            Width           =   1905
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nama Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   1380
            Width           =   1965
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   420
            Width           =   1920
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
         Height          =   4575
         Left            =   -74760
         TabIndex        =   15
         Top             =   1200
         Width           =   9615
         Begin VB.TextBox txtRateIndex 
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
            Left            =   5760
            TabIndex        =   8
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtNilaiIndexStandar 
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
            Left            =   3000
            TabIndex        =   7
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtKodeDetailKomp 
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
            Left            =   3000
            MaxLength       =   5
            TabIndex        =   5
            Top             =   840
            Width           =   4215
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
            Left            =   3000
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1320
            Width           =   6255
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   4320
            TabIndex        =   16
            Top             =   3960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dcKomponenIndex 
            Height          =   315
            Left            =   3000
            TabIndex        =   0
            Top             =   360
            Width           =   4215
            _ExtentX        =   7435
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
            Height          =   2055
            Left            =   240
            TabIndex        =   17
            Top             =   2280
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
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
            TabIndex        =   18
            Top             =   2280
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Rate Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4560
            TabIndex        =   23
            Top             =   1860
            Width           =   945
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Index Standar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   1800
            Width           =   1605
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Nama Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   420
            Width           =   1965
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Nama Detail Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   1380
            Width           =   2505
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Kode Detail Komponen Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   900
            Width           =   2445
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   54
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
      Picture         =   "frmDataKomponenIndex.frx":0D56
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8760
      Picture         =   "frmDataKomponenIndex.frx":3717
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataKomponenIndex.frx":449F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDataKomponenIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilter As String

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Call subDcSource
    cmdhapus.Enabled = False
    cmdUbah.Enabled = False

    Call loadDataGrid

    txtKodeJenisKomp.Enabled = True
    Me.sstDataPenunjang.Tab = 0
End Sub

Private Sub cmdBatal_Click()

    Select Case sstDataPenunjang.Tab
        Case 0
            txtKodeJenisKomp.Enabled = True
            clear
            cmdhapus.Enabled = False
            cmdUbah.Enabled = False
            cmdsimpan.Enabled = True

        Case 1
            dcJenisKomp.Enabled = True
            clear
            Text1.Text = ""
            cmdhapus.Enabled = False
            cmdUbah.Enabled = False
            cmdsimpan.Enabled = True

        Case 2
            dcKomponenIndex.Enabled = True
            clear
            cmdhapus.Enabled = False
            cmdUbah.Enabled = False
            cmdsimpan.Enabled = True

        Case 3
            dcJabatan.Enabled = True
            clear
            cmdhapus.Enabled = False
            cmdUbah.Enabled = False
            cmdsimpan.Enabled = True

        Case 4
            dcPendidikan.Enabled = True
            clear
            cmdhapus.Enabled = False
            cmdUbah.Enabled = False
            cmdsimpan.Enabled = True

    End Select

    Call subDcSource
    Call loadDataGrid
    Call clear
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan
    Select Case sstDataPenunjang.Tab
        Case 0
            Set rs = Nothing
            strSQL = "select * from JenisKomponenIndex where kdJenisKomponenIndex = '" & txtKodeJenisKomp & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "Kode Jenis Komponen Index Pegawai Sudah Ada"
                txtKodeJenisKomp.SetFocus
                Exit Sub
            End If

            Set rs = Nothing
            strSQL = "insert into JenisKomponenIndex(KdJenisKomponenIndex,JenisKomponenIndex) " & _
            "values ('" & txtKodeJenisKomp & "', '" & txtJenisKomp & "')"

            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 1
            Set rs = Nothing
            strSQL = "select * from V_DataKomponenIndex where KdKomponenIndex = '" & txtKodeKomp & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "Kode Jenis Komponen Index Pegawai Sudah Ada"
                dcJenisKomp.SetFocus
                Exit Sub
            End If

            Set rs = Nothing
            strSQL = "insert into KomponenIndex(KdKomponenIndex,KomponenIndex,KdJenisKomponenIndex) " & _
            "values ('" & txtKodeKomp & "','" & txtNamaKomponen & "','" & Text1 & "')"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs2 = Nothing

        Case 2
            Set rs = Nothing
            strSQL = "select * from V_DetailKomponenIndex where KdDetailKomponenIndex = '" & txtKodeDetailKomp & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "Kode Detail Komponen Index Sudah Ada"
                dcKomponenIndex.SetFocus
                Exit Sub
            End If

            Set rs = Nothing
            strSQL = "insert into DetailKomponenIndex (KdDetailKomponenIndex,KdKomponenIndex,DetailKomponenIndex,NilaiIndexStandar,RateIndex)" & _
            "values ('" & txtKodeDetailKomp & "', '" & Text2 & "','" & txtNamaDetailKomp & "','" & txtNilaiIndexStandar & "', '" & txtRateIndex & "')"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 3
            Set rs = Nothing
            strSQL = "select * from V_KonversiJabatanKeDetailKomponenIndex where NamaJabatan = '" & dcJabatan.Text & "' AND DetailKomponenIndex = '" & dcDetailKomponenIndexJabatan.Text & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "Konversi Jabatan ke Detail Komponen Index Sudah Ada"
                dcJabatan.SetFocus
                Exit Sub
            End If

            Set rs = Nothing
            strSQL = "insert into ConvertDetailKomponenIndexToJabatan (KdJabatan,KdDetailKomponenIndex)" & _
            "values ('" & dcJabatan.BoundText & "', '" & dcDetailKomponenIndexJabatan.BoundText & "')"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 4
            Set rs = Nothing
            strSQL = "select * from V_KonversiPendidikanKeDetailKomponenIndex where Pendidikan = '" & dcPendidikan.Text & "' AND DetailKomponenIndex = '" & dcDetailKomponenIndexPendidikan.Text & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "Konversi Pendidikan ke Detail Komponen Index Sudah Ada"
                dcPendidikan.SetFocus
                Exit Sub
            End If

            Set rs = Nothing
            strSQL = "insert into ConvertDetailKomponenIndexToPendidikan (KdPendidikan,KdDetailKomponenIndex)" & _
            "values ('" & dcPendidikan.BoundText & "', '" & dcDetailKomponenIndexPendidikan.BoundText & "')"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
    End Select
    Call cmdBatal_Click
    Exit Sub
errSimpan:
    msubPesanError
End Sub

Private Sub cmdUbah_Click()
    Select Case sstDataPenunjang.Tab
        Case 0
            Set rs = Nothing
            strSQL = "update JenisKomponenIndex set JenisKomponenIndex = '" & txtJenisKomp & "' where KdJenisKomponenIndex = '" & txtKodeJenisKomp & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 1
            Set rs = Nothing
            strSQL = "update KomponenIndex set KomponenIndex = '" & txtNamaKomponen & "' where KdKomponenIndex = '" & txtKodeKomp & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 2
            Set rs = Nothing
            strSQL = "update DetailKomponenIndex set DetailKomponenIndex = '" & txtNamaDetailKomp & "' where KdDetailKomponenIndex = '" & txtKodeDetailKomp & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 3
            Set rs = Nothing
            strSQL = "update ConvertDetailKomponenIndexToJabatan set KdDetailKomponenIndex = '" & dcDetailKomponenIndexJabatan.BoundText & "' where KdJabatan = '" & dcJabatan.BoundText & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 4
            Set rs = Nothing
            strSQL = "update ConvertDetailKomponenIndexToPendidikan set KdDetailKomponenIndex = '" & dcDetailKomponenIndexPendidikan.BoundText & "' where KdPendidikan = '" & dcPendidikan.BoundText & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

    End Select
    Call cmdBatal_Click
End Sub

Private Sub cmdHapus_Click()
    Select Case sstDataPenunjang.Tab
        Case 0
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

        Case 1
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

        Case 2
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

        Case 3
            Set rs = Nothing
            strSQL = "select * from ConvertDetailKomponenIndexToJabatan where KdJabatan = '" & dcJabatan.BoundText & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "data tidak dapat Di Hapus"
                dgJabatan.SetFocus
                Exit Sub
            End If

            Set rs = Nothing
            strSQL = "delete ConvertDetailKomponenIndexToJabatan where KdJabatan = '" & dcJabatan.BoundText & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 4
            Set rs = Nothing
            strSQL = "select * from ConvertDetailKomponenIndexToPendidikan where KdPendidikan = '" & dcPendidikan.BoundText & "'"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount <> 0 Then
                MsgBox "data tidak dapat Di Hapus"
                dgPendidikan.SetFocus
                Exit Sub
            End If

            Set rs = Nothing
            strSQL = "delete ConvertDetailKomponenIndexToPendidikan where KdPendidikan = '" & dcPendidikan.BoundText & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

    End Select
    Call cmdBatal_Click
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Sub clear()
    Select Case sstDataPenunjang.Tab
        Case 0
            txtKodeJenisKomp.Text = ""
            txtJenisKomp.Text = ""
            txtKodeJenisKomp.SetFocus

        Case 1
            txtKodeKomp.Text = ""
            txtNamaKomponen.Text = ""
            dcJenisKomp.Text = ""
            dcJenisKomp.SetFocus

        Case 2
            txtKodeDetailKomp.Text = ""
            txtNamaDetailKomp.Text = ""
            txtNilaiIndexStandar.Text = ""
            txtRateIndex.Text = ""
            dcKomponenIndex.Text = ""
            dcKomponenIndex.SetFocus

        Case 3
            dcJabatan.Text = ""
            dcDetailKomponenIndexJabatan.Text = ""
            txtNISJabatan.Text = ""
            txtRIJabatan.Text = ""
            dcJabatan.SetFocus

        Case 4
            dcPendidikan.Text = ""
            dcDetailKomponenIndexPendidikan.Text = ""
            txtNISPendidikan.Text = ""
            txtRIPendidikan.Text = ""
            dcPendidikan.SetFocus
    End Select
End Sub

Sub tampilData()
    Select Case sstDataPenunjang.Tab
        Case 0
            Set rs = Nothing
            strSQL = "select * from JenisKomponenIndex"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJenisKomp.DataSource = rs
            dgJenisKomp.Columns(0).DataField = rs(0).Name
            dgJenisKomp.Columns(1).DataField = rs(1).Name
            dgJenisKomp.ReBind
            Set rs = Nothing

        Case 1
            Set rs = Nothing
            strSQL = "select * from V_DataKomponenIndex "
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgKomponenIndex.DataSource = rs
            dgKomponenIndex.Columns(0).DataField = rs(0).Name
            dgKomponenIndex.Columns(1).DataField = rs(1).Name
            dgKomponenIndex.Columns(2).DataField = rs(2).Name
            dgKomponenIndex.Columns(3).DataField = rs(3).Name
            dgKomponenIndex.ReBind
            Set rs = Nothing

        Case 2
            Set rs = Nothing
            strSQL = "select * from V_DetailKomponenIndex"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgDetailKomponenIndex.DataSource = rs
            dgDetailKomponenIndex.Columns(0).DataField = rs(0).Name
            dgDetailKomponenIndex.Columns(1).DataField = rs(1).Name
            dgDetailKomponenIndex.Columns(2).DataField = rs(2).Name
            dgDetailKomponenIndex.Columns(3).DataField = rs(3).Name
            dgDetailKomponenIndex.Columns(4).DataField = rs(4).Name
            dgDetailKomponenIndex.Columns(5).DataField = rs(5).Name
            dgDetailKomponenIndex.ReBind
            Set rs = Nothing

        Case 3
            Set rs = Nothing
            strSQL = "select * from V_KonversiJabatanKeDetailKomponenIndex"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJabatan.DataSource = rs
            dgJabatan.Columns(0).DataField = rs(0).Name
            dgJabatan.Columns(1).DataField = rs(1).Name
            dgJabatan.Columns(2).DataField = rs(2).Name
            dgJabatan.Columns(3).DataField = rs(3).Name
            dgJabatan.Columns(4).DataField = rs(4).Name
            dgJabatan.Columns(5).DataField = rs(5).Name
            dgJabatan.ReBind
            Set rs = Nothing

        Case 4
            Set rs = Nothing
            strSQL = "select * from V_KonversiPendidikanKeDetailKomponenIndex"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgPendidikan.DataSource = rs
            dgPendidikan.Columns(0).DataField = rs(0).Name
            dgPendidikan.Columns(1).DataField = rs(1).Name
            dgPendidikan.Columns(2).DataField = rs(2).Name
            dgPendidikan.Columns(3).DataField = rs(3).Name
            dgPendidikan.Columns(4).DataField = rs(4).Name
            dgPendidikan.Columns(5).DataField = rs(5).Name
            dgPendidikan.ReBind
            Set rs = Nothing
    End Select
End Sub

Private Sub dgJenisKomp_Click()
    cmdsimpan.Enabled = False
    cmdUbah.Enabled = True
    cmdhapus.Enabled = True
    txtKodeJenisKomp.Text = dgJenisKomp.Columns(0).Value
    txtJenisKomp.Text = dgJenisKomp.Columns(1).Value
End Sub

Private Sub dgJenisKomp_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKodeJenisKomp.Text = dgJenisKomp.Columns(0).Value
    txtJenisKomp.Text = dgJenisKomp.Columns(1).Value
End Sub

Private Sub dgKomponenIndex_Click()
    cmdsimpan.Enabled = False
    cmdUbah.Enabled = True
    cmdhapus.Enabled = True
    Text1.Text = dgKomponenIndex.Columns(0).Value
    dcJenisKomp.Text = dgKomponenIndex.Columns(1).Value
    txtKodeKomp.Text = dgKomponenIndex.Columns(2).Value
    txtNamaKomponen.Text = dgKomponenIndex.Columns(3).Value
End Sub

Private Sub dgKomponenIndex_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    Text1.Text = dgKomponenIndex.Columns(0).Value
    dcJenisKomp.Text = dgKomponenIndex.Columns(1).Value
    txtKodeKomp.Text = dgKomponenIndex.Columns(2).Value
    txtNamaKomponen.Text = dgKomponenIndex.Columns(3).Value
End Sub

Private Sub dgDetailKomponenIndex_Click()
    cmdsimpan.Enabled = False
    cmdUbah.Enabled = True
    cmdhapus.Enabled = True
    Text2.Text = dgDetailKomponenIndex.Columns(0).Value
    dcKomponenIndex.Text = dgDetailKomponenIndex.Columns(1).Value
    txtKodeDetailKomp.Text = dgDetailKomponenIndex.Columns(2).Value
    txtNamaDetailKomp.Text = dgDetailKomponenIndex.Columns(3).Value
    txtNilaiIndexStandar.Text = dgDetailKomponenIndex.Columns(4).Value
    txtRateIndex.Text = dgDetailKomponenIndex.Columns(5).Value
End Sub

Private Sub dgDetailKomponenIndex_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    Text2.Text = dgDetailKomponenIndex.Columns(0).Value
    dcKomponenIndex.Text = dgDetailKomponenIndex.Columns(1).Value
    txtKodeDetailKomp.Text = dgDetailKomponenIndex.Columns(2).Value
    txtNamaDetailKomp.Text = dgDetailKomponenIndex.Columns(3).Value
    txtNilaiIndexStandar.Text = dgDetailKomponenIndex.Columns(4).Value
    txtRateIndex.Text = dgDetailKomponenIndex.Columns(5).Value
End Sub

Private Sub dgJabatan_Click()
    On Error Resume Next
    cmdsimpan.Enabled = False
    cmdUbah.Enabled = True
    cmdhapus.Enabled = True
    dcJabatan.Text = dgJabatan.Columns(1).Value
    dcDetailKomponenIndexJabatan.Text = dgJabatan.Columns(3).Value
    txtNISJabatan.Text = dgJabatan.Columns(4).Value
    txtRIJabatan.Text = dgJabatan.Columns(5).Value
End Sub

Private Sub dgJabatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    dcJabatan.Text = dgJabatan.Columns(1).Value
    dcDetailKomponenIndexJabatan.Text = dgJabatan.Columns(3).Value
    txtNISJabatan.Text = dgJabatan.Columns(4).Value
    txtRIJabatan.Text = dgJabatan.Columns(5).Value
End Sub

Private Sub dgPendidikan_Click()
    cmdsimpan.Enabled = False
    cmdUbah.Enabled = True
    cmdhapus.Enabled = True
    dcPendidikan.Text = dgPendidikan.Columns(1).Value
    dcDetailKomponenIndexPendidikan.Text = dgPendidikan.Columns(3).Value
    txtNISPendidikan.Text = dgPendidikan.Columns(4).Value
    txtRIPendidikan.Text = dgPendidikan.Columns(5).Value
End Sub

Private Sub dgPendidikan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    dcPendidikan.Text = dgPendidikan.Columns(1).Value
    dcDetailKomponenIndexPendidikan.Text = dgPendidikan.Columns(3).Value
    txtNISPendidikan.Text = dgPendidikan.Columns(4).Value
    txtRIPendidikan.Text = dgPendidikan.Columns(5).Value
End Sub

Sub SetComboJenisKomp()
    Set rs = Nothing
    rs.Open "Select * from JenisKomponenIndex ", dbConn, , adLockOptimistic
    Set dcJenisKomp.RowSource = rs
    dcJenisKomp.ListField = rs.Fields(1).Name
    Set rs = Nothing
End Sub

Sub SetComboNamaKomponenIndex()
    Set rs = Nothing
    rs.Open "Select * from KomponenIndex ", dbConn, , adLockOptimistic
    Set dcKomponenIndex.RowSource = rs
    dcKomponenIndex.ListField = rs.Fields(1).Name
    Set rs = Nothing
End Sub

Private Sub dcJenisKomp_Click(Area As Integer)
    strFilter = " WHERE JenisKomponenIndex like '%" & dcJenisKomp.Text & "%'"
    strSQL = "Select * from JenisKomponenIndex " & strFilter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgJenisKomp2.DataSource = rs
    Text1.Text = dgJenisKomp2.Columns(0)
End Sub

Private Sub dcKomponenIndex_Click(Area As Integer)
    strFilter = " WHERE KomponenIndex like '%" & dcKomponenIndex.Text & "%'"
    strSQL = "Select * from KomponenIndex " & strFilter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgKomponenIndex2.DataSource = rs
    Text2.Text = dgKomponenIndex2.Columns(0)
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call subDcSource
    Call loadDataGrid
End Sub

Private Sub txtKodeJenisKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtJenisKomp.SetFocus
End Sub

Private Sub dcJenisKomp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKodeKomp.SetFocus
End Sub

Private Sub txtKodeKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNamaKomponen.SetFocus
End Sub

Private Sub dcKomponenIndex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKodeDetailKomp.SetFocus
End Sub

Private Sub txtKodeDetailKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNamaDetailKomp.SetFocus
End Sub

Private Sub txtNamaDetailKomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNilaiIndexStandar.SetFocus
End Sub

Private Sub txtNilaiIndexStandar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtRateIndex.SetFocus
End Sub

Private Sub dcJabatan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcDetailKomponenIndexJabatan.SetFocus
End Sub

Private Sub dcPendidikan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcDetailKomponenIndexPendidikan.SetFocus
End Sub

Sub SetComboDetailKomponenIndexJabatan()
    Set rs = Nothing
    rs.Open "Select * from DetailKomponenIndex ", dbConn, , adLockOptimistic
    Set dcDetailKomponenIndexJabatan.RowSource = rs
    dcDetailKomponenIndexJabatan.ListField = rs.Fields(2).Name
    dcDetailKomponenIndexJabatan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Sub SetComboDetailKomponenIndexPendidikan()
    Set rs = Nothing
    rs.Open "Select * from DetailKomponenIndex WHERE KdKomponenIndex = 201 ", dbConn, , adLockOptimistic
    Set dcDetailKomponenIndexPendidikan.RowSource = rs
    dcDetailKomponenIndexPendidikan.ListField = rs.Fields(2).Name
    dcDetailKomponenIndexPendidikan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Sub SetComboJabatan()
    Set rs = Nothing
    rs.Open "Select * from Jabatan", dbConn, , adLockOptimistic
    Set dcJabatan.RowSource = rs
    dcJabatan.ListField = rs.Fields(1).Name
    dcJabatan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Sub SetComboPendidikan()
    Set rs = Nothing
    rs.Open "Select * from Pendidikan", dbConn, , adLockOptimistic
    Set dcPendidikan.RowSource = rs
    dcPendidikan.ListField = rs.Fields(1).Name
    dcPendidikan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Private Sub subDcSource()
    Call SetComboJenisKomp
    Call SetComboNamaKomponenIndex
    Call SetComboDetailKomponenIndexJabatan
    Call SetComboDetailKomponenIndexPendidikan
    Call SetComboJabatan
    Call SetComboPendidikan
End Sub

Private Sub loadDataGrid()
    On Error GoTo hell
    'case 0
    Set rs = Nothing
    strSQL = "select * from JenisKomponenIndex"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgJenisKomp.DataSource = rs
    dgJenisKomp.Columns(0).DataField = rs(0).Name
    dgJenisKomp.Columns(1).DataField = rs(1).Name
    dgJenisKomp.Columns(0).Caption = " KD JENIS KOMPONEN INDEX"
    dgJenisKomp.Columns(1).Caption = " JENIS KOMPONEN INDEX"
    dgJenisKomp.Columns(0).Width = 2500
    dgJenisKomp.Columns(1).Width = 6300
    Set rs = Nothing

    'case 1
    Set rs = Nothing
    strSQL = "select * from V_DataKomponenIndex"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgKomponenIndex.DataSource = rs
    dgKomponenIndex.Columns(0).DataField = rs(0).Name
    dgKomponenIndex.Columns(1).DataField = rs(1).Name
    dgKomponenIndex.Columns(2).DataField = rs(2).Name
    dgKomponenIndex.Columns(3).DataField = rs(3).Name
    dgKomponenIndex.Columns(0).Caption = "KD JENIS KOMPONEN INDEX"
    dgKomponenIndex.Columns(1).Caption = "JENIS KOMPONEN INDEX"
    dgKomponenIndex.Columns(2).Caption = "KD KOMPONEN INDEX"
    dgKomponenIndex.Columns(3).Caption = "KOMPONEN INDEX"
    dgKomponenIndex.Columns(0).Width = 0
    dgKomponenIndex.Columns(1).Width = 2500
    dgKomponenIndex.Columns(2).Width = 2000
    dgKomponenIndex.Columns(3).Width = 4000
    Set rs = Nothing

    'Case 2
    Set rs = Nothing
    strSQL = "select * from V_DetailKomponenIndex"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgDetailKomponenIndex.DataSource = rs
    dgDetailKomponenIndex.Columns(0).DataField = rs(0).Name
    dgDetailKomponenIndex.Columns(1).DataField = rs(1).Name
    dgDetailKomponenIndex.Columns(2).DataField = rs(2).Name
    dgDetailKomponenIndex.Columns(3).DataField = rs(3).Name
    dgDetailKomponenIndex.Columns(4).DataField = rs(4).Name
    dgDetailKomponenIndex.Columns(5).DataField = rs(5).Name
    dgDetailKomponenIndex.Columns(0).Caption = "KD KOMPONEN INDEX"
    dgDetailKomponenIndex.Columns(1).Caption = "KOMPONEN INDEX"
    dgDetailKomponenIndex.Columns(2).Caption = "KD DETAIL KOMP INDEX"
    dgDetailKomponenIndex.Columns(3).Caption = "DETAIL KOMP INDEX"
    dgDetailKomponenIndex.Columns(4).Caption = "NILAI INDEX STANDAR"
    dgDetailKomponenIndex.Columns(5).Caption = "RATE INDEX"
    dgDetailKomponenIndex.Columns(0).Width = 0
    dgDetailKomponenIndex.Columns(1).Width = 1500
    dgDetailKomponenIndex.Columns(2).Width = 2200
    dgDetailKomponenIndex.Columns(3).Width = 1750
    dgDetailKomponenIndex.Columns(4).Width = 2000
    dgDetailKomponenIndex.Columns(5).Width = 1300
    Set rs = Nothing

    'Case 3
    Set rs = Nothing
    strSQL = "select * from V_KonversiJabatanKeDetailKomponenIndex"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgJabatan.DataSource = rs
    dgJabatan.Columns(0).DataField = rs(0).Name
    dgJabatan.Columns(1).DataField = rs(1).Name
    dgJabatan.Columns(2).DataField = rs(2).Name
    dgJabatan.Columns(3).DataField = rs(3).Name
    dgJabatan.Columns(4).DataField = rs(4).Name
    dgJabatan.Columns(5).DataField = rs(5).Name
    dgJabatan.Columns(0).Caption = "KD JABATAN"
    dgJabatan.Columns(1).Caption = "JABATAN"
    dgJabatan.Columns(2).Caption = "KD DETAIL KOMPONEN INDEX"
    dgJabatan.Columns(3).Caption = "DETAIL KOMPONEN INDEX"
    dgJabatan.Columns(4).Caption = "NILAI INDEX STANDAR"
    dgJabatan.Columns(5).Caption = "RATE INDEX"
    dgJabatan.Columns(0).Width = 1200
    dgJabatan.Columns(1).Width = 1700
    dgJabatan.Columns(2).Width = 2500
    dgJabatan.Columns(3).Width = 2400
    dgJabatan.Columns(4).Width = 2000
    dgJabatan.Columns(5).Width = 1200
    Set rs = Nothing

    'Case 4
    Set rs = Nothing
    strSQL = "select * from V_KonversiPendidikanKeDetailKomponenIndex"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgPendidikan.DataSource = rs
    dgPendidikan.Columns(0).DataField = rs(0).Name
    dgPendidikan.Columns(1).DataField = rs(1).Name
    dgPendidikan.Columns(2).DataField = rs(2).Name
    dgPendidikan.Columns(3).DataField = rs(3).Name
    dgPendidikan.Columns(4).DataField = rs(4).Name
    dgPendidikan.Columns(5).DataField = rs(5).Name
    dgPendidikan.Columns(0).Caption = "KD PENDIDIKAN"
    dgPendidikan.Columns(1).Caption = "PENDIDIKAN"
    dgPendidikan.Columns(2).Caption = "KD DETAIL KOMPONEN INDEX"
    dgPendidikan.Columns(3).Caption = "DETAIL KOMPONEN INDEX"
    dgPendidikan.Columns(4).Caption = "NILAI INDEX STANDAR"
    dgPendidikan.Columns(5).Caption = "RATE INDEX"
    dgPendidikan.Columns(0).Width = 1400
    dgPendidikan.Columns(1).Width = 1700
    dgPendidikan.Columns(2).Width = 2500
    dgPendidikan.Columns(3).Width = 2400
    dgPendidikan.Columns(4).Width = 2000
    dgPendidikan.Columns(5).Width = 1200
    Set rs = Nothing
    Exit Sub
hell:
    Call msubPesanError
End Sub

