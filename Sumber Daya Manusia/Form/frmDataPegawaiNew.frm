VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataPegawaiNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Pegawai "
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataPegawaiNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   14940
   Begin MSComctlLib.ListView LVRuangan 
      Height          =   4215
      Left            =   11400
      TabIndex        =   98
      Top             =   1800
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Ruangan"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1695
      Left            =   3840
      TabIndex        =   83
      Top             =   5400
      Visible         =   0   'False
      Width           =   5175
      Begin MSDataListLib.DataCombo dcEselon 
         Height          =   315
         Left            =   2040
         TabIndex        =   84
         Top             =   840
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcTitle 
         Height          =   315
         Left            =   2040
         TabIndex        =   91
         Top             =   1200
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gelar"
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
         Index           =   13
         Left            =   720
         TabIndex        =   92
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eselon"
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
         Index           =   7
         Left            =   240
         TabIndex        =   86
         Top             =   840
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenjang Jabatan"
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
         Index           =   27
         Left            =   240
         TabIndex        =   85
         Top             =   480
         Visible         =   0   'False
         Width           =   1200
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   78
      Top             =   8505
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   26300
            Text            =   "Cetak (F1)"
            TextSave        =   "Cetak (F1)"
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
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   9360
      TabIndex        =   29
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpan 
      Appearance      =   0  'Flat
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   12240
      TabIndex        =   31
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdtutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   13680
      TabIndex        =   32
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   7680
      TabIndex        =   28
      Top             =   7920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   10800
      TabIndex        =   30
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdRiwayat 
      Appearance      =   0  'Flat
      Caption         =   "Riwayat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDetail 
      Appearance      =   0  'Flat
      Caption         =   "Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtParameter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   26
      Top             =   8040
      Width           =   2775
   End
   Begin VB.ComboBox cbJK 
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
      ItemData        =   "frmDataPegawaiNew.frx":0CCA
      Left            =   1680
      List            =   "frmDataPegawaiNew.frx":0CD4
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   62
      Top             =   4920
      Width           =   14775
      Begin VB.CommandButton cmdAlamat 
         Appearance      =   0  'Flat
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13560
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   960
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dgPegawai 
         Height          =   2535
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
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
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   13680
         TabIndex        =   63
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   36
      Top             =   960
      Width           =   14775
      Begin VB.CommandButton cmdRuangan 
         Caption         =   "..."
         Height          =   255
         Left            =   11280
         TabIndex        =   96
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdAutoNik 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   9240
         TabIndex        =   95
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtMasaKerja 
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
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   90
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtUmur 
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
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   89
         Top             =   3120
         Width           =   1335
      End
      Begin VB.ComboBox cbRhesus 
         Height          =   330
         Left            =   13560
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   3120
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dcKewarganegaraan 
         Height          =   330
         Left            =   11280
         TabIndex        =   79
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.CommandButton cmdLampiranFile 
         Caption         =   "Lampiran File"
         Height          =   375
         Left            =   12840
         TabIndex        =   77
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CommandButton cmdRiwayatStatus 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   9000
         TabIndex        =   76
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdRiwayatPangkat 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   14160
         TabIndex        =   74
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtGol 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   67
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtNegara 
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
         Left            =   11280
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtStatusResus 
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
         Left            =   13560
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtNamaPanggilan 
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
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtNamaKeluarga 
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
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtIDPegawai 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtNama 
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
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtTptLhr 
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
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txtNIP 
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
         Left            =   6480
         MaxLength       =   30
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo dcJnsPeg 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         ForeColor       =   0
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
      Begin MSDataListLib.DataCombo dcPangkat 
         Height          =   315
         Left            =   6480
         TabIndex        =   11
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcJabatan 
         Height          =   315
         Left            =   6480
         TabIndex        =   13
         Top             =   1320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcJurusan 
         Height          =   315
         Left            =   6480
         TabIndex        =   12
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSMask.MaskEdBox meTglLahir 
         Height          =   300
         Left            =   1560
         TabIndex        =   7
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcStatusPegawai 
         Height          =   315
         Left            =   6480
         TabIndex        =   16
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSMask.MaskEdBox meTglMasuk 
         Height          =   300
         Left            =   1560
         TabIndex        =   8
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meTglKeluar 
         Height          =   300
         Left            =   6480
         TabIndex        =   9
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         HideSelection   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcStatusKawin 
         Height          =   315
         Left            =   11280
         TabIndex        =   20
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcSuku 
         Height          =   315
         Left            =   11280
         TabIndex        =   21
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcAgama 
         Height          =   315
         Left            =   11280
         TabIndex        =   22
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcDarah 
         Height          =   315
         Left            =   11280
         TabIndex        =   23
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcTypePegawai 
         Height          =   315
         Left            =   6480
         TabIndex        =   14
         Top             =   1680
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcDetailKategoryPegawai 
         Height          =   315
         Left            =   6480
         TabIndex        =   15
         Top             =   2040
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcRuanganKerja 
         Height          =   315
         Left            =   11280
         TabIndex        =   17
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcIdPegawai 
         Height          =   315
         Left            =   11280
         TabIndex        =   18
         Top             =   960
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcJenjang 
         Height          =   315
         Left            =   6480
         TabIndex        =   94
         Top             =   2760
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Lainnya"
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
         Index           =   31
         Left            =   9840
         TabIndex        =   97
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kategory Pegawai"
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
         Index           =   30
         Left            =   4680
         TabIndex        =   93
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Umur"
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
         Index           =   29
         Left            =   3120
         TabIndex        =   88
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Masa Kerja"
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
         Index           =   28
         Left            =   3120
         TabIndex        =   87
         Top             =   3480
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Negara Asal"
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
         Left            =   9840
         TabIndex        =   80
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   2
         Left            =   8760
         TabIndex        =   75
         Top             =   2400
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   7
         Left            =   13920
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   6
         Left            =   9360
         TabIndex        =   71
         Top             =   1680
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   5
         Left            =   8760
         TabIndex        =   70
         Top             =   2400
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   3
         Left            =   11520
         TabIndex        =   69
         Top             =   3600
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label14 
         Caption         =   "data harus diisi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   68
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   1
         Left            =   2280
         TabIndex        =   66
         Top             =   2400
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   0
         Left            =   4320
         TabIndex        =   65
         Top             =   1320
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Index           =   4
         Left            =   4320
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pegawai Atasan"
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
         Index           =   26
         Left            =   9840
         TabIndex        =   61
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Kerja"
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
         Index           =   25
         Left            =   9840
         TabIndex        =   60
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kewarganegaraan"
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
         Index           =   24
         Left            =   9840
         TabIndex        =   59
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detail Kategori Pegawai"
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
         Index           =   23
         Left            =   4680
         TabIndex        =   58
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Rhesus"
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
         Index           =   22
         Left            =   12480
         TabIndex        =   57
         Top             =   3120
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Pegawai"
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
         Index           =   21
         Left            =   4680
         TabIndex        =   56
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gol. Darah"
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
         Index           =   20
         Left            =   9840
         TabIndex        =   55
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agama"
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
         Index           =   19
         Left            =   9840
         TabIndex        =   54
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suku"
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
         Index           =   18
         Left            =   9840
         TabIndex        =   53
         Top             =   2400
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Nikah"
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
         Index           =   17
         Left            =   9840
         TabIndex        =   52
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl.Keluar"
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
         Index           =   16
         Left            =   5520
         TabIndex        =   51
         Top             =   3480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Panggilan"
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
         Index           =   15
         Left            =   240
         TabIndex        =   50
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Keluarga"
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
         Index           =   14
         Left            =   240
         TabIndex        =   49
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl.Masuk"
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
         Index           =   12
         Left            =   240
         TabIndex        =   48
         Top             =   3480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Pegawai"
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
         Index           =   11
         Left            =   4680
         TabIndex        =   47
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID"
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
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   240
         TabIndex        =   44
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
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
         Index           =   3
         Left            =   240
         TabIndex        =   43
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Lahir"
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
         Index           =   4
         Left            =   240
         TabIndex        =   42
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl.Lahir"
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
         Index           =   5
         Left            =   240
         TabIndex        =   41
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pangkat Golongan"
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
         Index           =   6
         Left            =   4680
         TabIndex        =   40
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   8
         Left            =   4680
         TabIndex        =   39
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kualifikasi Pendidikan"
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
         Index           =   9
         Left            =   4680
         TabIndex        =   38
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIP / NRK"
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
         Index           =   10
         Left            =   4680
         TabIndex        =   37
         Top             =   240
         Width           =   705
      End
   End
   Begin MSDataListLib.DataCombo dcParamStatus 
      Height          =   315
      Left            =   4080
      TabIndex        =   27
      Top             =   8040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483644
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   81
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataPegawaiNew.frx":0CDE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmDataPegawaiNew.frx":233C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "Cari Nama "
      Height          =   255
      Left            =   240
      TabIndex        =   73
      Top             =   8040
      Width           =   1815
   End
End
Attribute VB_Name = "frmDataPegawaiNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msg As VbMsgBoxResult
Dim StatusUbah As String
Dim pesan As VbMsgBoxResult

Private Sub subLoadFormRiwayatPegawai()
    On Error GoTo errLoad
    mstrIdPegawai = dgPegawai.Columns(0).Value

    With frmRiwayatPegawai
        .Show
        .txtIdPegawai.Text = dgPegawai.Columns(0).Value
        .txtJenisPegawai.Text = frmDataPegawaiNew.dcJnsPeg.Text
        .txtNamaPegawai.Text = frmDataPegawaiNew.dgPegawai.Columns(3).Value
        .txtJabatan.Text = frmDataPegawaiNew.dgPegawai.Columns(14).Value
        If frmDataPegawaiNew.dgPegawai.Columns(6) = "L" Then
            .txtSex.Text = "Laki Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cbJK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTptLhr.SetFocus
End Sub

Private Sub cbJK_LostFocus()
    
    Call AutoNIP
End Sub

Private Sub cbRhesus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdSimpan.SetFocus
End Sub

Private Sub cmdAlamat_Click()
    If txtIdPegawai.Text = "" Then
        MsgBox "Pegawai belum dipilih  ", vbCritical, "Validasi"
    Else
        frmDataPegawaiNew.Enabled = False
        frmDataAlamatPegawai.Show
    End If
End Sub

Private Sub CmdBatal_Click()
    Call clearData
    Call subLoadDcSource
    Call subLoadGridSource
End Sub

Private Sub cmdCetak_Click()
    FrmCetakDataPegawai.Show
End Sub

Private Sub cmdDetail_Click()
    If txtIdPegawai.Text = "" Then
        MsgBox "Pegawai belum dipilih  ", vbCritical, "Validasi"
        Exit Sub
    End If
    If dcJabatan.BoundText = "" Then MsgBox "Silahkan lengkapi dan simpan jabatan pegawai", vbExclamation, "Validasi": Exit Sub
    strSQL = "Select * from V_S_Pegawai where IdPegawai = '" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then Exit Sub

    frmDataPegawaiNew.Enabled = False
    With frmDetailPegawai
        .Show
        .txtIdPegawai.Text = mstrIdPegawai
        .txtNamaLengkap.Text = txtNama.Text
        .txtJabatan.Text = dcJabatan.Text
    End With
End Sub

Private Sub cmdHapus_Click()
On Error GoTo xxx
    If txtIdPegawai.Text = "" Then Exit Sub
    If mblnAdmin = False Then
        MsgBox "User tidak bisa menghapus Data Pegawai.", vbInformation
        Exit Sub
    End If
    If MsgBox("Hapus data pegawai dengan nomor " & txtIdPegawai.Text & "", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    Set rs = Nothing
    strSQL = "Select distinct * from V_RiwayatDataPegawai where IdPegawai = '" & txtIdPegawai.Text & "'"
    Call msubRecFO(rs, strSQL)
    
    If rs.EOF = False Then
        MsgBox "Hapus terlebih dahulu semua riwayat Pegawai.", vbCritical
        cmdRiwayat.SetFocus
        Exit Sub
    Else
        dbConn.Execute "DELETE DataCurrentPegawai where IdPegawai = '" & txtIdPegawai.Text & "'"
        dbConn.Execute "DELETE DataPegawai where IdPegawai = '" & txtIdPegawai.Text & "'"
        MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
        Call subLoadGridSource
        Call CmdBatal_Click
    End If
    
Exit Sub
xxx:
    MsgBox "Riwayat data digunakan, tidak dapat dihapus ", vbOKOnly, "Validasi"
End Sub

Private Sub cmdLampiranFile_Click()

Const QUOTE As String = """"
Dim Path As String

    Dim fso As New FileSystemObject
    If txtIdPegawai.Text = "" Then Exit Sub
    
    If fso.FolderExists(mstrPathFileSDM & "\SDM_" & txtIdPegawai.Text) = False Then fso.CreateFolder mstrPathFileSDM & "\SDM_" & txtIdPegawai.Text

    Path = Replace(mstrPathFileSDM & "\SDM_" & txtIdPegawai.Text, QUOTE, QUOTE & QUOTE)
    Shell "explorer.exe /e, " & Path, vbNormalFocus
End Sub

Private Sub cmdRiwayat_Click()
    If txtIdPegawai.Text = "" Then
        MsgBox "Pegawai belum dipilih  ", vbCritical, "Validasi"
    Else

'        strSQL = "Select * from V_S_Pegawai where IdPegawai = '" & mstrIdPegawai & "'"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF = True Then Exit Sub
        If dcJabatan.Text = "" Then
            MsgBox "Data jabatan harus di isi terlebih dahulu", vbInformation
            dcJabatan.SetFocus
        Exit Sub
        End If
        
        Call subLoadFormRiwayatPegawai
        frmDataPegawaiNew.Enabled = False
    End If
End Sub

Private Sub cmdRiwayatRuangan_Click()

End Sub

Private Sub cmdRuangan_Click()
Dim ii As Integer
    If LVRuangan.Visible = False Then
        LVRuangan.Visible = True
    Else
        LVRuangan.Visible = False
    End If
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim sStatus As String
    
    If dcPangkat.Text <> "" Then
        If Periksa("datacombo", dcPangkat, "Pangkat Golongan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcJurusan.Text <> "" Then
        If Periksa("datacombo", dcJurusan, "Kualifikasi Pendidikan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcJabatan.Text <> "" Then
        If Periksa("datacombo", dcJabatan, "Jabatan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcTypePegawai.Text <> "" Then
        If Periksa("datacombo", dcTypePegawai, "Tipe Pegawai Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcDetailKategoryPegawai.Text <> "" Then
        If Periksa("datacombo", dcDetailKategoryPegawai, "Detail Kategory Pegawai Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcStatusPegawai.Text <> "" Then
        If Periksa("datacombo", dcStatusPegawai, "Status Pegawai Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcRuanganKerja.Text <> "" Then
        If Periksa("datacombo", dcRuanganKerja, "Ruang Kerja Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcIdPegawai.Text <> "" Then
        If Periksa("datacombo", dcIdPegawai, "Pegawai Atasan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcKewarganegaraan.Text <> "" Then
        If Periksa("datacombo", dcKewarganegaraan, "Kewarganegaraan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcStatusKawin.Text <> "" Then
        If Periksa("datacombo", dcStatusKawin, "Status Nikah Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcSuku.Text <> "" Then
        If Periksa("datacombo", dcSuku, "Suku Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcAgama.Text <> "" Then
        If Periksa("datacombo", dcAgama, "Agama Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If Periksa("datacombo", dcJnsPeg, "Jenis pegawai kosong") = False Then Exit Sub
    If Periksa("text", txtNama, "Nama pegawai kosong") = False Then Exit Sub
    If Periksa("text", cbJK, "Jenis kelamin kosong") = False Then Exit Sub
    If Periksa("datacombo", dcStatusPegawai, "Status Pegawai Kosong") = False Then Exit Sub
    If StatusUbah = "U" Then
        pesan = MsgBox("Yakin akan mengubah data pegawai " & strNama & " ", vbQuestion + vbYesNo, "Konfirmasi")
        If pesan = vbYes Then
            If sp_DataPegawai = False Then Exit Sub
            If sp_DataCurrentPegawai = False Then Exit Sub

            MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"

        End If
    End If
    If StatusUbah = "N" Then

        If sp_DataPegawai = False Then Exit Sub
        If sp_DataCurrentPegawai = False Then Exit Sub
        
        strSQL = "delete from RuanganLainnya where idpegawai='" & txtIdPegawai.Text & "'"
        Call msubRecFO(rs, strSQL)
        For i = 1 To LVRuangan.ListItems.Count
            If LVRuangan.ListItems(i).Checked = True Then
                strSQL = "insert into RuanganLainnya values ('" & txtIdPegawai.Text & "','" & Right(LVRuangan.ListItems(i).key, 3) & "')"
                Call msubRecFO(rs, strSQL)
            End If
        Next
        

        MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"

    End If
    Call CmdBatal_Click
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAgama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcAgama.BoundText = ""
End Sub

Private Sub dcAgama_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcDarah.SetFocus

If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcAgama.Text)) = 0 Then dcDarah.SetFocus: Exit Sub
        If dcAgama.MatchedWithList = True Then dcDarah.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select kdagama,agama from agama where agama LIKE '%" & dcAgama.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcAgama.BoundText = dbRst(0).Value
        dcAgama.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcDarah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcDarah.BoundText = ""
End Sub

Private Sub dcDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbRhesus.SetFocus
End Sub

Private Sub dcDetailKategoryPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcDetailKategoryPegawai.BoundText = ""
End Sub

Private Sub dcDetailKategoryPegawai_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcStatusPegawai.SetFocus

If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcDetailKategoryPegawai.Text)) = 0 Then dcStatusPegawai.SetFocus: Exit Sub
        If dcDetailKategoryPegawai.MatchedWithList = True Then dcStatusPegawai.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select kddetailkategorypegawai,detailkategorypegawai from DetailKategoryPegawai where detailkategorypegawai LIKE '%" & dcDetailKategoryPegawai.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcDetailKategoryPegawai.BoundText = dbRst(0).Value
        dcDetailKategoryPegawai.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcEselon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcEselon.BoundText = ""
End Sub

Private Sub dcEselon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcTypePegawai.SetFocus
End Sub

Private Sub dcIdPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcIdPegawai.BoundText = ""
End Sub

Private Sub dcIdPegawai_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcKewarganegaraan.SetFocus

If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcIdPegawai.Text)) = 0 Then dcKewarganegaraan.SetFocus: Exit Sub
        If dcIdPegawai.MatchedWithList = True Then dcKewarganegaraan.SetFocus: Exit Sub
        'Call msubRecFO(dbRst, "select kdStatus,Status from StatusPegawai where Status LIKE '%" & dcIdPegawai.Text & "%' ")
        Call msubRecFO(dbRst, "SELECT TOP (100) PERCENT dbo.DataPegawai.IdPegawai, dbo.DataPegawai.NamaLengkap FROM dbo.DataPegawai INNER JOIN dbo.DataCurrentPegawai ON dbo.DataPegawai.IdPegawai = dbo.DataCurrentPegawai.IdPegawai WHERE dbo.DataPegawai.NamaLengkap LIKE '%" & dcIdPegawai.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcIdPegawai.BoundText = dbRst(0).Value
        dcIdPegawai.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJabatan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcJabatan.BoundText = ""
End Sub

Private Sub dcJabatan_KeyPress(KeyAscii As Integer)
'    On Error Resume Next
'    If KeyAscii = 13 Then
'        strSQL = "select KdJenisJabatan from Jabatan where KdJabatan='" & dcJabatan.BoundText & "'"
'        Call msubRecFO(rs, strSQL)
'        If rs(0).Value = "04" Then
'            dcJenjang.Enabled = True
'            dcJenjang.SetFocus
'        Else
'            dcJenjang.Enabled = False
'            dcEselon.SetFocus
'        End If
'        If rs(0).Value = "01" Then
'            dcEselon.Enabled = True
'            dcEselon.SetFocus
'        Else
'            dcEselon.Enabled = False
'            dcTypePegawai.SetFocus
'        End If
'    End If
    
On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
    If rs.EOF = True Then
        If dcJenjang.Enabled = False Then
            dcTypePegawai.SetFocus
        Else
            dcEselon.SetFocus
        Exit Sub
        End If
    Else

        If Len(Trim(dcJabatan.Text)) = 0 Then
            If rs(0).Value = "04" Then
                dcJenjang.Enabled = True
                dcJenjang.SetFocus
            Else
                dcJenjang.Enabled = False
                dcEselon.SetFocus
            End If
            If rs(0).Value = "01" Then
                dcEselon.Enabled = True
                dcEselon.SetFocus
            Else
                dcEselon.Enabled = False
                dcTypePegawai.SetFocus
            End If
            Exit Sub
        End If
        If dcJabatan.MatchedWithList = True Then
            If rs(0).Value = "04" Then
                dcJenjang.Enabled = True
                dcJenjang.SetFocus
            Else
                dcJenjang.Enabled = False
                dcEselon.SetFocus
            End If
            If rs(0).Value = "01" Then
                dcEselon.Enabled = True
                dcEselon.SetFocus
            Else
                dcEselon.Enabled = False
                dcTypePegawai.SetFocus
            End If
            Exit Sub
        End If
    End If
    Call msubRecFO(dbRst, "select KdJabatan, NamaJabatan from Jabatan where NamaJabatan LIKE '%" & dcJabatan.Text & "%' ")
    If dbRst.EOF = True Then Exit Sub
    dcJabatan.BoundText = dbRst(0).Value
    dcJabatan.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJabatan_LostFocus()
    On Error Resume Next
    strSQL = "select KdJenisJabatan from Jabatan where KdJabatan='" & dcJabatan.BoundText & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then Exit Sub
    If rs(0).Value = "04" Then
        dcJenjang.Enabled = True
        dcJenjang.SetFocus
    Else
        dcJenjang.Enabled = False
        dcEselon.SetFocus
    End If
    If rs(0).Value = "01" Then
        dcEselon.Enabled = True
        dcEselon.SetFocus
    Else
        dcEselon.Enabled = False
        dcTypePegawai.SetFocus
    End If
End Sub

Private Sub dcJenjang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcJenjang.BoundText = ""
End Sub

Private Sub dcJenjang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcEselon.Enabled = True Then
            dcEselon.SetFocus
        ElseIf dcEselon.Enabled = False Then
            dcTypePegawai.SetFocus
        End If
    End If
End Sub

Private Sub dcJnsPeg_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcTitle.SetFocus

On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcJnsPeg.Text)) = 0 Then dcTitle.SetFocus: Exit Sub
        If dcJnsPeg.MatchedWithList = True Then dcTitle.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select kdJenisPegawai,jenispegawai from jenispegawai where JenisPegawai LIKE '%" & dcJnsPeg.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcJnsPeg.BoundText = dbRst(0).Value
        dcJnsPeg.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJurusan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcJurusan.BoundText = ""
End Sub

Private Sub dcJurusan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcJabatan.SetFocus

On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcJurusan.Text)) = 0 Then dcJabatan.SetFocus: Exit Sub
        If dcJurusan.MatchedWithList = True Then dcJabatan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select kdkualifikasijurusan,kualifikasijurusan from KualifikasiJurusan where kualifikasijurusan LIKE '%" & dcJurusan.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcJurusan.BoundText = dbRst(0).Value
        dcJurusan.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKewarganegaraan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtNegara.SetFocus

If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKewarganegaraan.Text)) = 0 Then txtNegara.SetFocus: Exit Sub
        If dcKewarganegaraan.MatchedWithList = True Then txtNegara.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select KdNegara,NamaNegara from Negara where NamaNegara LIKE '%" & dcKewarganegaraan.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcKewarganegaraan.BoundText = dbRst(0).Value
        dcKewarganegaraan.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPangkat_Change()
    On Error Resume Next
    strSQL = "select dbo.GolonganPegawai.NamaGolongan from dbo.Pangkat INNER JOIN dbo.GolonganPegawai ON dbo.Pangkat.KdGolongan = dbo.GolonganPegawai.KdGolongan where dbo.Pangkat.KdPangkat = '" & dcPangkat.BoundText & "'"
    Call msubRecFO(rs, strSQL)
    txtGol.Text = rs(0).Value
    
    If dcPangkat.BoundText = "" Then
        txtGol.Text = ""
    End If
End Sub

Private Sub dcPangkat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcPangkat.BoundText = "": txtGol.Text = ""
End Sub

Private Sub dcPangkat_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcJurusan.SetFocus

On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcPangkat.Text)) = 0 Then dcJurusan.SetFocus: Exit Sub
        If dcPangkat.MatchedWithList = True Then dcJurusan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select kdPangkat,NamaPangkat from Pangkat WHERE NamaPangkat LIKE '%" & dcPangkat.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcPangkat.BoundText = dbRst(0).Value
        dcPangkat.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcParamStatus_Click(Area As Integer)
    Call subLoadGridSource
End Sub

Private Sub dcRuanganKerja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcRuanganKerja.BoundText = ""
End Sub

Private Sub dcRuanganKerja_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcIdPegawai.SetFocus

If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcRuanganKerja.Text)) = 0 Then dcIdPegawai.SetFocus: Exit Sub
        If dcRuanganKerja.MatchedWithList = True Then dcIdPegawai.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select KdRuangan,NamaRuangan from Ruangan where NamaRuangan LIKE '%" & dcRuanganKerja.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcRuanganKerja.BoundText = dbRst(0).Value
        dcRuanganKerja.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcStatusKawin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcStatusKawin.BoundText = ""
End Sub

Private Sub dcStatusKawin_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then dcSuku.SetFocus

If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcStatusKawin.Text)) = 0 Then dcSuku.SetFocus: Exit Sub
        If dcStatusKawin.MatchedWithList = True Then dcSuku.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select kdstatusperkawinan,statusperkawinan from statusperkawinan where statusperkawinan LIKE '%" & dcStatusKawin.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcStatusKawin.BoundText = dbRst(0).Value
        dcStatusKawin.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcStatusPegawai_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcRuanganKerja.SetFocus

If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcStatusPegawai.Text)) = 0 Then dcRuanganKerja.SetFocus: Exit Sub
        If dcStatusPegawai.MatchedWithList = True Then dcRuanganKerja.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select kdStatus,Status from StatusPegawai where Status LIKE '%" & dcStatusPegawai.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcStatusPegawai.BoundText = dbRst(0).Value
        dcStatusPegawai.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcSuku_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcSuku.BoundText = ""
End Sub

Private Sub dcSuku_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcAgama.SetFocus

If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcSuku.Text)) = 0 Then dcAgama.SetFocus: Exit Sub
        If dcSuku.MatchedWithList = True Then dcAgama.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select kdsuku,suku from suku where suku LIKE '%" & dcSuku.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcSuku.BoundText = dbRst(0).Value
        dcSuku.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcTitle_Change()
'    If dcTitle.Text = "Tn." Then
'        cbJK.Text = "L"
'        cbJK.Enabled = False
'    ElseIf dcTitle.Text = "Ny." Or dcTitle.Text = "Nn." Then
'        cbJK.Text = "P"
'        cbJK.Enabled = False
'    Else
'        cbJK.Enabled = True
'    End If
End Sub

Private Sub dcTitle_Click(Area As Integer)
    dcTitle_Change
End Sub

Private Sub dcTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcTitle.BoundText = ""
End Sub

Private Sub dcTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNama.SetFocus
End Sub

Private Sub dcTypePegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then dcTypePegawai.BoundText = ""
End Sub

Private Sub dcTypePegawai_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dcDetailKategoryPegawai.SetFocus

If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcTypePegawai.Text)) = 0 Then dcDetailKategoryPegawai.SetFocus: Exit Sub
        If dcTypePegawai.MatchedWithList = True Then dcDetailKategoryPegawai.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select kdtypepegawai,typepegawai from TypePegawai where typepegawai LIKE '%" & dcTypePegawai.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcTypePegawai.BoundText = dbRst(0).Value
        dcTypePegawai.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPegawai
    WheelHook.WheelHook dgPegawai
End Sub

Private Sub dgPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & dgPegawai.Bookmark & "/" & dgPegawai.ApproxCount
    With dgPegawai
    '//yayang.agus 2014-08-07
        txtIdPegawai.Text = IIf(.Columns(0).Text = "", "", .Columns(0).Text)
        dcJnsPeg.Text = IIf(.Columns(1).Text = "", "", .Columns(1).Text)
        dcTitle.Text = IIf(.Columns(2).Text = "", "", .Columns(2).Text)
        txtNama.Text = IIf(.Columns(3).Text = "", "", .Columns(3).Text)
        txtNamaKeluarga.Text = IIf(.Columns(4).Text = "", "", .Columns(4).Text)
        txtNamaPanggilan.Text = IIf(.Columns(5).Text = "", "", .Columns(5).Text)
        cbJK.Text = IIf(.Columns(6).Text = "", "", .Columns(6).Text)
        txtTptLhr.Text = IIf(.Columns(7).Text = "", "", .Columns(7).Text)
        meTglLahir.Text = IIf(.Columns(8).Text = "", "__/__/____", Format(.Columns(8).Value, "dd/mm/yyyy"))
        meTglMasuk.Text = IIf(.Columns(9).Text = "", "__/__/____", Format(.Columns(9).Value, "dd/mm/yyyy"))
        meTglKeluar.Text = IIf(.Columns(10).Text = "", "__/__/____", Format(.Columns(10).Value, "dd/mm/yyyy"))
        txtNIP.Text = IIf(.Columns(15).Text = "", "", .Columns(15).Text)
        dcPangkat.Text = IIf(.Columns(11).Text = "", "", .Columns(11).Text)
        txtGol.Text = IIf(.Columns(12).Text = "", "", .Columns(12).Text)
        dcJurusan.Text = IIf(.Columns(48).Text = "", "", .Columns(48).Text)
        dcJabatan.Text = IIf(.Columns(14).Text = "", "", .Columns(14).Text)
        dcJenjang.Text = IIf(.Columns(44).Text = "", "", .Columns(44).Text)
        dcEselon.Text = IIf(.Columns(46).Text = "", "", .Columns(46).Text)
        dcTypePegawai.Text = IIf(.Columns(22).Text = "", "", .Columns(22).Text)
        dcDetailKategoryPegawai.Text = IIf(.Columns(23).Text = "", "", .Columns(23).Text)
        dcStatusPegawai.Text = IIf(.Columns(16).Text = "", "", .Columns(16).Text)
        dcRuanganKerja.Text = IIf(.Columns(26).Text = "", "", .Columns(26).Text)
        dcIdPegawai.Text = IIf(.Columns(27).Text = "", "", .Columns(27).Text)
        If .Columns(24).Text = "" Then
            dcKewarganegaraan.Text = ""
        Else
            dcKewarganegaraan.Text = IIf(.Columns(24).Text = "1", "WNA", "WNI")
        End If
        txtNegara.Text = IIf(.Columns(25).Text = "", "", .Columns(25).Text)
        dcStatusKawin.Text = IIf(.Columns(17).Text = "", "", .Columns(17).Text)
        dcSuku.Text = IIf(.Columns(18).Text = "", "", .Columns(18).Text)
        dcAgama.Text = IIf(.Columns(19).Text = "", "", .Columns(19).Text)
        dcDarah.Text = IIf(.Columns(20).Text = "", "", .Columns(20).Text)
        cbRhesus.Text = IIf(.Columns(21).Text = "", "", .Columns(21).Text)
        
    '//
'        txtIDPegawai.Text = .Columns(0).Value
'        dcJnsPeg.Text = .Columns(1).Value
'        If IsNull(.Columns(2).Value) Then dcTitle.Text = "" Else dcTitle.Text = .Columns(2).Value
'        txtNama.Text = .Columns(3).Value
'        If IsNull(.Columns(4).Value) Then txtNamaKeluarga.Text = "" Else txtNamaKeluarga.Text = .Columns(4).Value
'        If IsNull(.Columns(3).Value) Then txtNamaPanggilan.Text = "" Else txtNamaPanggilan.Text = .Columns(3).Value
'        cbJK.Text = .Columns(6).Value
'        If IsNull(.Columns(7).Value) Then txtTptLhr.Text = "" Else txtTptLhr.Text = .Columns(7).Value
'        If IsNull(.Columns(8).Value) Then meTglLahir.Text = "__/__/____" Else meTglLahir.Text = .Columns(8).Value
'        If IsNull(.Columns(9).Value) Then meTglMasuk.Text = "__/__/____" Else meTglMasuk.Text = .Columns(9).Value
'        If IsNull(.Columns(10).Value) Then meTglKeluar.Text = "__/__/____" Else meTglKeluar.Text = .Columns(10).Value
'
'        If IsNull(.Columns(15).Value) Then txtNIP.Text = "" Else txtNIP.Text = .Columns(15).Value
'        'If IsNull(.Columns(21).Value) Then txtStatusResus.Text = "" Else txtStatusResus.Text = .Columns(21).Value
'        If IsNull(.Columns(21).Value) Then cbRhesus.Text = "" Else cbRhesus.Text = .Columns(21).Value
'        If IsNull(.Columns(24).Value) Then dcKewarganegaraan.BoundText = "" Else dcKewarganegaraan.BoundText = .Columns(24).Value
'        If IsNull(.Columns(25).Value) Then txtNegara.Text = "" Else txtNegara.Text = .Columns(25).Value
'
'        'If IsNull(.Columns(30).Value) Then dcPangkat.BoundText = "" Else dcPangkat.BoundText = .Columns(30).Value
'        If IsNull(.Columns(29).Value) Then dcPangkat.BoundText = "" Else dcPangkat.BoundText = .Columns(29).Value
'
'        'If IsNull(.Columns(32).Value) Then dcJurusan.BoundText = "" Else dcJurusan.BoundText = .Columns(32).Value
'        If IsNull(.Columns(31).Value) Then dcJurusan.BoundText = "" Else dcJurusan.BoundText = .Columns(31).Value
'
'        'If IsNull(.Columns(33).Value) Then dcJabatan.BoundText = "" Else dcJabatan.BoundText = .Columns(33).Value
'        If IsNull(.Columns(32).Value) Then dcJabatan.BoundText = "" Else dcJabatan.BoundText = .Columns(32).Value
'
'        'dcStatusPegawai.BoundText = .Columns(34).Value
'        dcStatusPegawai.BoundText = .Columns(33).Value
'
'        'If IsNull(.Columns(35).Value) Then dcStatusKawin.BoundText = "" Else dcStatusKawin.BoundText = .Columns(35).Value
'        If IsNull(.Columns(34).Value) Then dcStatusKawin.BoundText = "" Else dcStatusKawin.BoundText = .Columns(34).Value
'
'        'If IsNull(.Columns(36).Value) Then dcSuku.BoundText = "" Else dcSuku.BoundText = .Columns(36).Value
'        If IsNull(.Columns(35).Value) Then dcSuku.BoundText = "" Else dcSuku.BoundText = .Columns(35).Value
'
'        'If IsNull(.Columns(37).Value) Then dcAgama.BoundText = "" Else dcAgama.BoundText = .Columns(37).Value
'        If IsNull(.Columns(36).Value) Then dcAgama.BoundText = "" Else dcAgama.BoundText = .Columns(36).Value
'
'        'If IsNull(.Columns(38).Value) Then dcDarah.BoundText = "" Else dcDarah.BoundText = .Columns(38).Value
'        If IsNull(.Columns(37).Value) Then dcDarah.BoundText = "" Else dcDarah.BoundText = .Columns(37).Value
'
'        'If IsNull(.Columns(39).Value) Then dcTypePegawai.BoundText = "" Else dcTypePegawai.BoundText = .Columns(39).Value
'        If IsNull(.Columns(38).Value) Then dcTypePegawai.BoundText = "" Else dcTypePegawai.BoundText = .Columns(38).Value
'
'        'If IsNull(.Columns(40).Value) Then dcDetailKategoryPegawai.BoundText = "" Else dcDetailKategoryPegawai.BoundText = .Columns(40).Value
'        If IsNull(.Columns(39).Value) Then dcDetailKategoryPegawai.BoundText = "" Else dcDetailKategoryPegawai.BoundText = .Columns(39).Value
'
'        'If IsNull(.Columns(41).Value) Then dcRuanganKerja.BoundText = "" Else dcRuanganKerja.BoundText = .Columns(41).Value
'        If IsNull(.Columns(40).Value) Then dcRuanganKerja.BoundText = "" Else dcRuanganKerja.BoundText = .Columns(40).Value
'
'        If IsNull(.Columns(42).Value) Then dcIdPegawai.BoundText = "" Else dcIdPegawai.BoundText = .Columns(42).Value
'        If IsNull(.Columns(43).Value) Then dcJenjang.BoundText = "" Else dcJenjang.BoundText = .Columns(43).Value
'        If IsNull(.Columns(45).Value) Then dcEselon.BoundText = "" Else dcEselon.BoundText = .Columns(45).Value
'        If IsNull(.Columns(12).Value) Then txtGol.Text = "" Else txtGol.Text = .Columns(12).Value

        mstrIdPegawai = txtIdPegawai.Text
        strNama = .Columns(3).Value
    End With
    
    Dim ii As Integer
    LVRuangan.Visible = False
    For ii = 1 To LVRuangan.ListItems.Count
        LVRuangan.ListItems(ii).Checked = False
    Next
    strSQL = "select * from RuanganLainnya where idpegawai='" & txtIdPegawai.Text & "'"
    Call msubRecFO(rs, strSQL)
    For i = 0 To rs.RecordCount - 1
        For ii = 1 To LVRuangan.ListItems.Count
            If Right(LVRuangan.ListItems(ii).key, 3) = rs(1) Then
                LVRuangan.ListItems(ii).Checked = True
                Exit For
'            Else
'                LVRuangan.ListItems(ii).Checked = False
            End If
        Next
        rs.MoveNext
    Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    centerForm Me, MDIUtama
    txtNegara.Text = "INDONESIA"
    Call CmdBatal_Click
    
    cbRhesus.AddItem "+"
    cbRhesus.AddItem "-"
End Sub

Sub clearData()
    On Error Resume Next
    txtIdPegawai.Text = ""
    dcJnsPeg.BoundText = ""
    dcTitle.BoundText = ""
    txtNama.Text = ""
    txtNamaKeluarga.Text = ""
    txtNamaPanggilan.Text = ""
    cbJK.Text = ""
    txtTptLhr.Text = ""
    meTglLahir.Text = "__/__/____"
    meTglMasuk.Text = "__/__/____"
    meTglKeluar.Text = "__/__/____"
    txtNIP.Text = ""
    dcPangkat.BoundText = ""
    txtGol.Text = ""
    dcJurusan.BoundText = ""
    dcJabatan.BoundText = ""
    dcJenjang.BoundText = ""
    dcEselon.BoundText = ""
    dcTypePegawai.BoundText = ""
    dcDetailKategoryPegawai.BoundText = ""
    dcStatusPegawai.BoundText = ""
    dcRuanganKerja.BoundText = ""
    dcIdPegawai.BoundText = ""
    dcKewarganegaraan.BoundText = ""
    txtNegara.Text = "INDONESIA"
    dcStatusKawin.BoundText = ""
    dcSuku.BoundText = ""
    dcAgama.BoundText = ""
    dcDarah.BoundText = ""
    txtStatusResus.Text = ""
    cbRhesus.Text = ""
    
    
    
    
    StatusUbah = "N"
    CmdBatal.Caption = "&Batal"
    CmdSimpan.Caption = "&Simpan"
End Sub

Private Sub subLoadDcSource()
    strSQL = "select kdJenisPegawai,jenispegawai from jenispegawai where StatusEnabled='1' order by jenispegawai "
    Call msubDcSource(dcJnsPeg, rs, strSQL)

    strSQL = "select kdtitle,namatitle from Title Where StatusEnabled='1' order by namatitle "
    Call msubDcSource(dcTitle, rs, strSQL)

    strSQL = "select KdNegara,NamaNegara from Negara order by NamaNegara"
    Call msubDcSource(dcKewarganegaraan, rs, strSQL)

    strSQL = "select kdPangkat,NamaPangkat from Pangkat where StatusEnabled='1' order by namapangkat"
    Call msubDcSource(dcPangkat, rs, strSQL)

    strSQL = "select KdEselon,NamaEselon from Eselon where StatusEnabled='1' order by NamaEselon"
    Call msubDcSource(dcEselon, rs, strSQL)

    strSQL = "select kdkualifikasijurusan,kualifikasijurusan from KualifikasiJurusan where statusenabled='1' order by kualifikasijurusan"
    Call msubDcSource(dcJurusan, rs, strSQL)

    strSQL = "select kdJabatan,Namajabatan from jabatan where StatusEnabled='1' order by namajabatan"
    Call msubDcSource(dcJabatan, rs, strSQL)

    strSQL = "select kdStatus,Status from StatusPegawai where StatusEnabled='1' order by Status"
    Call msubDcSource(dcStatusPegawai, rs, strSQL)

    strSQL = "select kdStatus,Status from StatusPegawai where StatusEnabled='1' order by Status"
    Call msubDcSource(dcParamStatus, rs, strSQL)
    dcParamStatus.BoundText = rs(0).Value

    strSQL = "select kdtypepegawai,typepegawai from TypePegawai where StatusEnabled='1'  order by TypePegawai"
    Call msubDcSource(dcTypePegawai, rs, strSQL)

    strSQL = "select kddetailkategorypegawai,detailkategorypegawai from DetailKategoryPegawai where StatusEnabled='1'  order by detailkategorypegawai"
    Call msubDcSource(dcDetailKategoryPegawai, rs, strSQL)

    strSQL = "select KdRuangan,NamaRuangan from Ruangan where StatusEnabled='1' order by NamaRuangan"
    Call msubDcSource(dcRuanganKerja, rs, strSQL)

    strSQL = "SELECT TOP (100) PERCENT dbo.DataPegawai.IdPegawai, dbo.DataPegawai.NamaLengkap " & _
    "FROM dbo.DataPegawai INNER JOIN " & _
    "dbo.DataCurrentPegawai ON dbo.DataPegawai.IdPegawai = dbo.DataCurrentPegawai.IdPegawai " & _
    "WHERE (dbo.DataCurrentPegawai.KdStatus = '01') AND (dbo.DataPegawai.IdPegawai <> '8888888888') " & _
    "ORDER BY dbo.DataPegawai.NamaLengkap "
    Call msubDcSource(dcIdPegawai, rs, strSQL)

    strSQL = "select kdstatusperkawinan,statusperkawinan from statusperkawinan where StatusEnabled='1' order by statusperkawinan"
    Call msubDcSource(dcStatusKawin, rs, strSQL)

    strSQL = "select kdsuku,suku from suku where StatusEnabled='1' order by suku"
    Call msubDcSource(dcSuku, rs, strSQL)

    strSQL = "select kdagama,agama from agama where StatusEnabled='1' order by agama"
    Call msubDcSource(dcAgama, rs, strSQL)

    strSQL = "select kdgolongandarah,golongandarah from golongandarah where StatusEnabled='1' order by golongandarah"
    Call msubDcSource(dcDarah, rs, strSQL)

    strSQL = "select kdjenjang,namajenjangjabatan from jenjangjabatan where StatusEnabled='1' order by namajenjangjabatan"
    Call msubDcSource(dcJenjang, rs, strSQL)
    
    strSQL = "select * from Ruangan  where KdInstalasi in ('01','02','03')and StatusEnabled='1' order by NamaRuangan"
    Call msubRecFO(rs, strSQL)

    LVRuangan.ListItems.clear
    While Not rs.EOF
        LVRuangan.ListItems.add , "A" & rs(0).Value, rs(1).Value
        LVRuangan.ListItems("A" & rs(0)).Checked = False
        LVRuangan.ListItems("A" & rs(0)).ForeColor = vbBlack
        rs.MoveNext
    Wend
    LVRuangan.Sorted = True
End Sub

Private Sub meTglKeluar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNIP.SetFocus
End Sub

Private Sub meTglKeluar_LostFocus()
    On Error GoTo errTglLahir
    If meTglKeluar.Text = "__/__/____" Then Exit Sub
    If funcCekValidasiTgl("TglKeluar", meTglKeluar) <> "NoErr" Then Exit Sub
    Exit Sub
errTglLahir:
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
    & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Sub meTglLahir_Change()
On Error Resume Next
    Dim Tgl As Date
    If meTglLahir.Text = "__/__/____" Then Exit Sub
    Tgl = meTglLahir.Text
    txtUmur.Text = CStr(DateDiff("D", Tgl, Now()) \ 356) & " thn"
End Sub

Private Sub meTglLahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meTglMasuk.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo hell
    Dim pesan As VbMsgBoxResult
    If KeyCode = vbKeyF1 Then
        pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
        vLaporan = ""
        If pesan = vbYes Then vLaporan = "Print" Else vLaporan = "view"
        FrmCetakDataPegawai.Show
    End If

    Exit Sub
hell:
End Sub

Private Sub meTglLahir_LostFocus()
    On Error GoTo errTglLahir
    If meTglLahir.Text = "__/__/____" Then Exit Sub
    If funcCekValidasiTgl("TglLahir", meTglLahir) <> "NoErr" Then Exit Sub
    
    Call AutoNIP
    
    Exit Sub
errTglLahir:
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
    & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Sub meTglMasuk_Change()
On Error Resume Next
    Dim Tgl As Date
    If meTglMasuk.Text = "__/__/____" Then Exit Sub
    Tgl = meTglMasuk.Text
    txtMasaKErja.Text = CStr(DateDiff("D", Tgl, Now()) \ 356) & " thn " & CStr((DateDiff("D", Tgl, Now()) Mod 356) \ 30) & " bln"
    

End Sub

Private Sub meTglMasuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNIP.SetFocus
End Sub

Private Sub meTglMasuk_LostFocus()
    On Error GoTo errgoTglLahir
    If meTglMasuk.Text = "__/__/____" Then Exit Sub
    If funcCekValidasiTgl("TglMasuk", meTglMasuk) <> "NoErr" Then Exit Sub
    
    Call AutoNIP
    
    Exit Sub
errgoTglLahir:
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
    & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Sub AutoNIP()
On Error Resume Next

    Dim jmlorg As Integer
    Dim jnskelaminKode As String
    
    If cbJK.Text = "L" Then jnskelaminKode = "01"
    If cbJK.Text = "P" Then jnskelaminKode = "02"
    
    strSQL = "select COUNT(idpegawai) from DataPegawai where datepart(YEAR, TglMasuk ) ='" & Right(meTglMasuk.Text, 4) & "' and DATEPART(MONTH,tglmasuk)='" & Mid(meTglMasuk.Text, 4, 2) & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        jmlorg = CDbl(rs(0)) + 1
    Else
        jmlorg = 1
    End If
    txtNIP.Text = Right(meTglLahir.Text, 4) & Mid(meTglLahir.Text, 4, 2) & Left(meTglLahir.Text, 2) & Right(meTglMasuk.Text, 4) & Mid(meTglMasuk.Text, 4, 2) & Format(jmlorg, "0#") & jnskelaminKode
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub txtIDPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaKeluarga.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNamaKeluarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPanggilan.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNamaPanggilan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTptLhr.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNegara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcStatusKawin.SetFocus
End Sub

Private Sub txtNIP_GotFocus()
    
    Call AutoNIP
End Sub

Private Sub txtNIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPangkat.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtParameter_Change()
    Call subLoadGridSource
    strCetak = " where NamaLengkap LIKE '%" & txtParameter.Text & "%' order by NamaLengkap"
End Sub

Private Sub txtStatusResus_Change()
    txtStatusResus.MaxLength = 1
End Sub

Private Sub txtStatusResus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdSimpan.SetFocus
End Sub

Private Sub txtTptLhr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meTglLahir.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Sub subLoadGridSource()
    On Error GoTo hell
    Set rs = Nothing
    strSQL = "select * from V_M_DataPegawaiNew where [Nama Lengkap] LIKE '%" & txtParameter.Text & "%' and KdStatus = '" & dcParamStatus.BoundText & "' and IdPegawai <> '8888888888' order by [Nama Lengkap] "
    Call msubRecFO(rs, strSQL)
    Set dgPegawai.DataSource = rs
    lblJumData.Caption = "Data " & dgPegawai.Bookmark & "/" & dgPegawai.ApproxCount
    Call SetDataGrid
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub SetDataGrid()
    On Error Resume Next
    Dim i As Integer
    With dgPegawai
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i

        .Columns(0).Width = 1000
        .Columns(1).Width = 1300
        .Columns(2).Width = 800
        .Columns(2).Caption = "Gelar"
        .Columns(3).Width = 1800
        .Columns(6).Width = 300
        .Columns(7).Width = 1300
        .Columns(8).Width = 1000
        .Columns(9).Width = 1000
        .Columns(11).Width = 1500
        .Columns(12).Width = 1000
        .Columns(14).Width = 1500
        .Columns(15).Width = 1500
        .Columns(16).Width = 1500
        .Columns(17).Width = 1500
        .Columns(18).Width = 1500
        .Columns(19).Width = 1500
        .Columns(20).Width = 1500
        .Columns(21).Width = 1500
        .Columns(22).Width = 1500
        .Columns(23).Width = 1500
        .Columns(24).Width = 1500
        .Columns(25).Width = 1500
        .Columns(26).Width = 1500

    End With
End Sub

Private Function sp_DataPegawai() As Boolean
    sp_DataPegawai = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIdPegawai.Text)
        .Parameters.Append .CreateParameter("KdJenisPegawai", adChar, adParamInput, 3, IIf(dcJnsPeg.BoundText = "", Null, dcJnsPeg.BoundText))
        .Parameters.Append .CreateParameter("KdTitle", adChar, adParamInput, 2, IIf(dcTitle.Text = "", Null, dcTitle.BoundText))
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, Trim(txtNama.Text))
        .Parameters.Append .CreateParameter("NamaKeluarga", adVarChar, adParamInput, 50, IIf(txtNamaKeluarga.Text = "", Null, Trim(txtNamaKeluarga.Text)))
        .Parameters.Append .CreateParameter("NamaPanggilan", adVarChar, adParamInput, 50, IIf(txtNamaPanggilan.Text = "", Null, Trim(txtNamaPanggilan.Text)))
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, cbJK.Text)
        .Parameters.Append .CreateParameter("TempatLahir", adVarChar, adParamInput, 50, IIf(txtTptLhr.Text = "", Null, Trim(txtTptLhr.Text)))
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , IIf(meTglLahir = "__/__/____", Null, Format(meTglLahir, "dd/MM/yyyy")))
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , IIf(meTglMasuk = "__/__/____", Null, Format(meTglMasuk, "dd/MM/yyyy")))
        .Parameters.Append .CreateParameter("TglKeluar", adDate, adParamInput, , IIf(meTglKeluar = "__/__/____", Null, Format(meTglKeluar, "dd/MM/yyyy")))
        .Parameters.Append .CreateParameter("OutKode", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("statusCode", adChar, adParamInput, 1, "A")

        .ActiveConnection = dbConn
        .CommandText = "AUD_DataPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data Pegawai", vbCritical, "Validasi"
            sp_DataPegawai = False
        Else
            If Not IsNull(.Parameters("OutKode").Value) Then txtIdPegawai = .Parameters("OutKode").Value
            mstrIdPegawai = txtIdPegawai.Text

        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
End Function

Private Function sp_DataCurrentPegawai() As Boolean
    sp_DataCurrentPegawai = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIdPegawai.Text)
        .Parameters.Append .CreateParameter("KdPangkat", adVarChar, adParamInput, 2, IIf(dcPangkat.Text = "", Null, dcPangkat.BoundText))
        .Parameters.Append .CreateParameter("KdJabatan", adVarChar, adParamInput, 5, IIf(dcJabatan.Text = "", Null, dcJabatan.BoundText))
        .Parameters.Append .CreateParameter("KdEselon", adVarChar, adParamInput, 2, IIf(dcEselon.Text = "", Null, dcEselon.BoundText))
        .Parameters.Append .CreateParameter("NIP", adVarChar, adParamInput, 30, IIf(txtNIP.Text = "", Null, txtNIP.Text))
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, dcStatusPegawai.BoundText)
        
        .Parameters.Append .CreateParameter("KdKualifikasiJurusan", adVarChar, adParamInput, 4, IIf(dcJurusan.Text = "", Null, dcJurusan.BoundText))
        .Parameters.Append .CreateParameter("KdStatusPerkawinan", adChar, adParamInput, 2, IIf(dcStatusKawin.Text = "", Null, Trim(dcStatusKawin.BoundText)))
        .Parameters.Append .CreateParameter("KdSuku", adChar, adParamInput, 2, IIf(dcSuku.Text = "", Null, dcSuku.BoundText))
        .Parameters.Append .CreateParameter("KdAgama", adChar, adParamInput, 2, IIf(dcAgama.Text = "", Null, dcAgama.BoundText))
        .Parameters.Append .CreateParameter("KdGolonganDarah", adChar, adParamInput, 2, IIf(dcDarah.Text = "", Null, dcDarah.BoundText))
        '.Parameters.Append .CreateParameter("StatusRhesus", adChar, adParamInput, 1, IIf(txtStatusResus.Text = "", Null, txtStatusResus.Text))
        .Parameters.Append .CreateParameter("StatusRhesus", adChar, adParamInput, 1, IIf(cbRhesus.Text = "", Null, cbRhesus.Text))
        .Parameters.Append .CreateParameter("KdTypePegawai", adChar, adParamInput, 2, IIf(dcTypePegawai.Text = "", Null, dcTypePegawai.BoundText))
        .Parameters.Append .CreateParameter("KdDetailKategoryPegawai", adVarChar, adParamInput, 2, IIf(dcDetailKategoryPegawai.Text = "", Null, dcDetailKategoryPegawai.BoundText))
        .Parameters.Append .CreateParameter("KdNegara", adTinyInt, adParamInput, , IIf(dcKewarganegaraan.Text = "", Null, dcKewarganegaraan.BoundText))
        .Parameters.Append .CreateParameter("Negara", adVarChar, adParamInput, 50, IIf(txtNegara.Text = "", Null, txtNegara.Text))
        .Parameters.Append .CreateParameter("KdRuanganKerja", adChar, adParamInput, 3, IIf(dcRuanganKerja.Text = "", Null, dcRuanganKerja.BoundText))
        .Parameters.Append .CreateParameter("KdPegawaiAtasan", adChar, adParamInput, 10, IIf(dcIdPegawai.Text = "", Null, dcIdPegawai.BoundText))
        .Parameters.Append .CreateParameter("KdJenjang", adVarChar, adParamInput, 5, IIf(dcJenjang.Text = "", Null, dcJenjang.BoundText))
        .Parameters.Append .CreateParameter("PathFoto", adVarChar, adParamInput, 250, Null)
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, "A")

        .ActiveConnection = dbConn
        .CommandText = "AUD_DataCurrentPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Pegawai", vbCritical, "Validasi"
            sp_DataCurrentPegawai = False
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
End Function
