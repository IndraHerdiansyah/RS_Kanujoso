VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRiwayatUsulan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Tempat Bertugas Pegawai"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   Icon            =   "frmRiwayatUsulan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   9735
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
      Height          =   1095
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   9495
      Begin VB.TextBox txtNamaPegawai 
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
         Left            =   240
         MaxLength       =   50
         TabIndex        =   33
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtjk 
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
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   32
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txttempatbertugas 
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
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   31
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtjabatan 
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
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   30
         Top             =   600
         Width           =   2655
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
         TabIndex        =   37
         Top             =   360
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
         Left            =   2880
         TabIndex        =   36
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Bertugas"
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
         Left            =   3960
         TabIndex        =   35
         Top             =   360
         Width           =   1335
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
         Left            =   6240
         TabIndex        =   34
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.TextBox txtIdPegawai 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Height          =   330
      Left            =   0
      MaxLength       =   100
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox txtNoRiwayat 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Height          =   330
      Left            =   0
      MaxLength       =   100
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
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
      Left            =   6840
      TabIndex        =   8
      Top             =   8160
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
      Left            =   5400
      TabIndex        =   7
      Top             =   8160
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
      Left            =   8280
      TabIndex        =   9
      Top             =   8160
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
      Left            =   3960
      TabIndex        =   6
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Tempat Bertugas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   9495
      Begin VB.TextBox txtTTD 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   6840
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtSatuanKerja 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   3240
         MaxLength       =   75
         TabIndex        =   21
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         Height          =   330
         Left            =   4560
         MaxLength       =   100
         TabIndex        =   5
         Top             =   2160
         Width           =   4695
      End
      Begin VB.CheckBox chkTglAkhir 
         Caption         =   "Tgl Akhir Berlaku"
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
         Left            =   2400
         TabIndex        =   17
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtNoSuratKeputusan 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   4560
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1320
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpTglMulai 
         Height          =   330
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy"
         Format          =   18808832
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   330
         Left            =   2400
         TabIndex        =   3
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy"
         Format          =   18808832
         CurrentDate     =   38448
      End
      Begin MSDataListLib.DataCombo dcKdRuangan 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
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
      Begin MSDataListLib.DataCombo dcKdJabatan 
         Height          =   315
         Left            =   6240
         TabIndex        =   1
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
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
      Begin MSDataListLib.DataCombo dcEselon 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSComCtl2.DTPicker dtpTglSK 
         Height          =   330
         Left            =   2400
         TabIndex        =   23
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy"
         Format          =   18808832
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tanda Tangan SK"
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
         Left            =   6840
         TabIndex        =   26
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tgl SK"
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
         Left            =   2400
         TabIndex        =   24
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Satuan Kerja"
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
         TabIndex        =   22
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
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
         Left            =   4560
         TabIndex        =   18
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "No. SK"
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
         Left            =   4560
         TabIndex        =   15
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nama Ruangan"
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
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Mulai Berlaku"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label Label20 
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
         Left            =   6240
         TabIndex        =   12
         Top             =   360
         Width           =   585
      End
   End
   Begin MSDataGridLib.DataGrid dgTempatBertugas 
      Height          =   2775
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
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
      Left            =   7920
      Picture         =   "frmRiwayatUsulan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatUsulan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatUsulan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatUsulan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''update by splakuk 2010/8/14
Option Explicit

Private Sub chkTglAkhir_Click()
    If chkTglAkhir.Value = vbChecked Then dtpTglAkhir.Enabled = True Else dtpTglAkhir.Enabled = False
End Sub

Private Sub chkTglAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub cmdBatal_Click()
    Call subClearData
    dcKdRuangan.SetFocus
End Sub

Private Sub cmdHapus_Click()
On Error GoTo errHapus
    If dcKdJabatan.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatTempatBertugas WHERE NoRiwayat='" & txtnoriwayat.Text & "' "
    dbConn.Execute strSQL
    If sp_Riwayat("D") = False Then Exit Sub
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Call subLoadTempatBertugas
    Call subClearData
Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    If Periksa("datacombo", dcKdRuangan, "Nama Ruangan harus diisi ") = False Then Exit Sub
    If Periksa("datacombo", dcKdJabatan, "Nama Jabatan harus diisi ") = False Then Exit Sub
    
    If sp_Riwayat("A") = False Then Exit Sub
    If sp_RiwayatTempatBertugas = False Then Exit Sub

    Call subLoadTempatBertugas
    Call cmdBatal_Click
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Call frmRiwayatPegawai.subLoadRiwayatTempatBertugas
    Unload Me
End Sub

Private Sub dcEselon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglSK.SetFocus
End Sub

Private Sub dgTempatBertugas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
If dgTempatBertugas.ApproxCount = 0 Then Exit Sub
    With dgTempatBertugas
        dcKdRuangan.BoundText = .Columns(14).Value
        If IsNull(.Columns(4)) Then txtSatuanKerja.Text = "" Else txtSatuanKerja.Text = .Columns(4).Value
        dcKdJabatan.BoundText = .Columns(15).Value
        If IsNull(.Columns(16)) Then dcEselon.BoundText = "" Else dcEselon.BoundText = .Columns(16).Value
        dtpTglSK.Value = .Columns(7).Value
        If IsNull(.Columns(8)) Then txtNoSuratKeputusan.Text = "" Else txtNoSuratKeputusan.Text = .Columns(8).Value
        If IsNull(.Columns(9)) Then txtTTD.Text = "" Else txtTTD.Text = .Columns(9).Value
        dtpTglMulai.Value = dgTempatBertugas.Columns(10).Value
        If Len(Trim(.Columns(11).Value)) = 0 Then
            chkTglAkhir.Value = vbUnchecked
        Else
            chkTglAkhir.Value = vbChecked
            dtpTglAkhir.Value = .Columns(11).Value
        End If
        If IsNull(.Columns(12)) Then txtKeterangan.Text = "" Else txtKeterangan.Text = .Columns(12).Value
        txtnoriwayat.Text = .Columns(0).Value
        txtIDPegawai.Text = .Columns(1).Value
    End With
End Sub

Private Sub dtpTglAkhir_Change()
    dtpTglAkhir.MaxDate = Now
End Sub

Private Sub dtpTglMulai_Change()
    dtpTglMulai.MaxDate = Now
End Sub

Private Sub dtpTglSK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNoSuratKeputusan.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call SetComboRuangan
    Call SetComboJabatan
    Call SetComboEselon
    txtIDPegawai.Text = mstrIdPegawai
    Call subLoadTempatBertugas
End Sub

Private Sub subLoadTempatBertugas()
On Error GoTo errLoad
    strSQL = "SELECT * FROM V_RiwayatTempatbertugas where IdPegawai = '" & mstrIdPegawai & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgTempatBertugas.DataSource = rs
    With dgTempatBertugas
    
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
    
        .Columns(3).Width = 2000 'ruangan
        .Columns(4).Width = 2000 'satuan kerja
        .Columns(5).Width = 2000 'jabatan
        .Columns(6).Width = 1000 'eselon
        .Columns(7).Width = 1000 'tglsk
        .Columns(8).Width = 1500 'nosk
        .Columns(9).Width = 1500 'tandatanGan sk
        .Columns(10).Width = 1000 'tglmulai berlaku
        .Columns(11).Width = 1000 'tglakhirberlaku
        .Columns(12).Width = 2000 'keteranga
    End With
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subClearData()
On Error Resume Next
    txtnoriwayat.Text = ""
    txtIDPegawai.Text = ""
    dcKdRuangan.Text = ""
    dcKdJabatan.Text = ""
    dcEselon.Text = ""
    dtpTglSK.Value = Format(Now, "dd/mmmm/yyyy")
    dtpTglMulai.Value = Format(Now, "dd/mmmm/yyyy")
    dtpTglAkhir.Value = Format(Now, "dd/mmmm/yyyy")
    dtpTglAkhir.Enabled = False
    chkTglAkhir.Value = vbUnchecked
    txtTTD.Text = ""
    txtSatuanKerja.Text = ""
    txtNoSuratKeputusan.Text = ""
    txtKeterangan.Text = ""
End Sub

Private Sub dcKdRuangan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtSatuanKerja.SetFocus
End Sub

Private Sub dcKdJabatan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcEselon.SetFocus
End Sub

Private Sub dtpTglMulai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then chkTglAkhir.SetFocus
End Sub

Private Sub dtpTglAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatTempatBertugas
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtNoSuratKeputusan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtTTD.SetFocus
End Sub

 Sub SetComboRuangan()
 On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from Ruangan order by NamaRuangan ASC", dbConn, , adLockOptimistic
    Set dcKdRuangan.RowSource = rs
    dcKdRuangan.ListField = rs.Fields(1).Name
    dcKdRuangan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
Exit Sub
hell:
    Call msubPesanError
End Sub

Sub SetComboJabatan()
On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from Jabatan order by NamaJabatan ASC", dbConn, , adLockOptimistic
    Set dcKdJabatan.RowSource = rs
    dcKdJabatan.ListField = rs.Fields(1).Name
    dcKdJabatan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
Exit Sub
hell:
    Call msubPesanError
End Sub

Sub SetComboEselon()
On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from Eselon order by NamaEselon ASC", dbConn, , adLockOptimistic
    Set dcEselon.RowSource = rs
    dcEselon.ListField = rs.Fields(1).Name
    dcEselon.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtSatuanKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKdJabatan.SetFocus
End Sub

Private Sub txtTTD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglMulai.SetFocus
End Sub

Private Function sp_Riwayat(f_Status) As Boolean
On Error GoTo hell
    sp_Riwayat = True
    Set dbcmd = New ADODB.Command
    With dbcmd
    
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        If txtnoriwayat = "" Then
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtnoriwayat.Text)
        End If
        
        .Parameters.Append .CreateParameter("TglRiwayat", adDate, adParamInput, , Format(Now, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .Parameters.Append .CreateParameter("OutputNoRiwayat", adChar, adParamOutput, 10, Null)
                
                        
        .ActiveConnection = dbConn
        .CommandText = "AUD_Riwayat"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data nomor riwayat", vbCritical, "Validasi"
            sp_Riwayat = False
        Else
            If Not IsNull(.Parameters("Status").Value) Then txtnoriwayat.Text = .Parameters("OutputNoRiwayat").Value
        End If
        
        
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_RiwayatTempatBertugas() As Boolean
On Error GoTo hell
    sp_RiwayatTempatBertugas = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtnoriwayat.Text))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, Trim(txtIDPegawai.Text))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcKdRuangan.BoundText)
        .Parameters.Append .CreateParameter("SatuanKerja", adVarChar, adParamInput, 75, IIf(txtSatuanKerja.Text = "", Null, Trim(txtSatuanKerja.Text)))
        .Parameters.Append .CreateParameter("KdJabatan", adVarChar, adParamInput, 5, dcKdJabatan.BoundText)
        .Parameters.Append .CreateParameter("KdEselon", adVarChar, adParamInput, 2, IIf(dcEselon.Text = "", Null, dcEselon.BoundText))
        .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Format(dtpTglSK.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("NoSK", adVarChar, adParamInput, 30, IIf(txtNoSuratKeputusan.Text = "", Null, Trim(txtNoSuratKeputusan.Text)))
        .Parameters.Append .CreateParameter("TandaTanganSK", adVarChar, adParamInput, 50, IIf(txtTTD.Text = "", Null, Trim(txtTTD.Text)))
        .Parameters.Append .CreateParameter("TglMulaiBerlaku", adDate, adParamInput, , Format(dtpTglMulai.Value, "yyyy/MM/dd"))
        If chkTglAkhir.Value = vbChecked Then
            .Parameters.Append .CreateParameter("TglAkhirBerlaku", adDate, adParamInput, , Format(dtpTglAkhir.Value, "yyyy/MM/dd"))
        Else
            .Parameters.Append .CreateParameter("TglAkhirBerlaku", adDate, adParamInput, , Null)
        End If
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RTempatBertugas"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
Exit Function
hell:
    Call msubPesanError
End Function
