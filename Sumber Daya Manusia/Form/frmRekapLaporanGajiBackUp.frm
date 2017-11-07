VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRekapLaporanGaji 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   14595
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   14535
      Begin TabDlg.SSTab SSTab1 
         Height          =   3375
         Left            =   120
         TabIndex        =   17
         Top             =   3840
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   5953
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Komponen Gaji"
         TabPicture(0)   =   "frmRekapLaporanGajiBackUp.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dpPembayaranGaji"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Potongan Gaji"
         TabPicture(1)   =   "frmRekapLaporanGajiBackUp.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgPotonganGaji"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid dpPembayaranGaji 
            Height          =   2895
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   5106
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
         Begin MSDataGridLib.DataGrid dgPotonganGaji 
            Height          =   2895
            Left            =   -74880
            TabIndex        =   18
            Top             =   360
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   5106
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
      End
      Begin MSDataGridLib.DataGrid dgRekapLaporan 
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   4471
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
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   7200
         Width           =   14295
         Begin VB.CommandButton cmdTutup 
            Caption         =   "Tutup"
            Height          =   375
            Left            =   11760
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cetak Rekap"
            Height          =   375
            Left            =   9840
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Cetak Struk Gaji (F1)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5400
            TabIndex        =   14
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   14295
         Begin VB.TextBox txtNamaPegawai 
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   240
            MaxLength       =   100
            TabIndex        =   16
            Top             =   480
            Width           =   2775
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cari"
            Height          =   375
            Left            =   7920
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
         Begin VB.Frame Frame4 
            Caption         =   "Periode"
            Height          =   735
            Left            =   9120
            TabIndex        =   5
            Top             =   120
            Width           =   4935
            Begin MSComCtl2.DTPicker dtpTglAwal 
               Height          =   330
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   2295
               _ExtentX        =   4048
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
               Format          =   129761280
               UpDown          =   -1  'True
               CurrentDate     =   38448
            End
            Begin MSComCtl2.DTPicker dtpTglAhir 
               Height          =   330
               Left            =   2520
               TabIndex        =   7
               Top             =   240
               Width           =   2295
               _ExtentX        =   4048
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
               Format          =   129761280
               UpDown          =   -1  'True
               CurrentDate     =   38448
            End
         End
         Begin MSDataListLib.DataCombo dcJabatan 
            Height          =   315
            Left            =   3240
            TabIndex        =   8
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcPangkat 
            Height          =   315
            Left            =   5280
            TabIndex        =   9
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label1 
            Caption         =   "Nama Pegawai"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Jabatan"
            Height          =   255
            Left            =   3240
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Pangkat"
            Height          =   255
            Left            =   5280
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
      End
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12720
      Picture         =   "frmRekapLaporanGajiBackUp.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRekapLaporanGajiBackUp.frx":0DC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmRekapLaporanGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgRekapLaporan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call subLoadGridTab
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLodGrid
    Call subLoadGridTab
    Call subLoadDcSource
End Sub

Public Sub subLoadDcSource()
    On Error GoTo errLoad
    Call msubDcSource(dcJabatan, rs, "SELECT KdJabatan, NamaJabatan FROM Jabatan where NamaJabatan LIKE '%" & dcJabatan.Text & "%' and StatusEnabled = 1 ORDER BY NamaJabatan")
    If rs.EOF = False Then dcJabatan.BoundText = rs(0).Value
    Call msubDcSource(dcPangkat, rs, "SELECT KdPangkat, NamaPangkat FROM Pangkat where NamaPangkat LIKE '%" & dcPangkat.Text & "%' and StatusEnabled = 1 ORDER BY NamaPangkat")
    If rs.EOF = False Then dcPangkat.BoundText = rs(0).Value
    Exit Sub
errLoad:
End Sub

Public Sub subLodGrid()
    On Error GoTo bawah
    Set rs = Nothing
    strSQL = "SELECT * FROM V_RekapLaporanGajiPegawai where [Nama Lengkap] LIKE '%" & txtNamaPegawai.Text & "%' ORDER BY [Nama Lengkap]"
    Call msubRecFO(rs, strSQL)
    Set dgRekapLaporan.DataSource = rs
bawah:
End Sub

Public Sub subLoadGridTab()
    On Error GoTo bawah
    Select Case SSTab1.Tab
        Case 0
            Set rs = Nothing
            strSQL = "SELECT*FROM V_PembayaranGajiPegawai where NamaLengkap='" & dgRekapLaporan.Columns("Nama Lengkap") & "' and tglPembayaran = '" & Format(dgRekapLaporan.Columns("tglPembayaran"), "yyyy/MM/dd") & "'"
            Call msubRecFO(rs, strSQL)
            Set dpPembayaranGaji.DataSource = rs
        Case 1
            Set rs = Nothing
            strSQL = "SELECT*FROM V_PembayaranPotonganGajiPegawai where NamaLengkap='" & dgRekapLaporan.Columns("Nama Lengkap") & "'and tglPembayaran = '" & Format(dgRekapLaporan.Columns("tglPembayaran"), "yyyy/MM/dd") & "'"
            Call msubRecFO(rs, strSQL)
            Set dgPotonganGaji.DataSource = rs
    End Select
bawah:
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call subLoadGridTab
End Sub

Private Sub txtNamaPegawai_Change()
    Call subLodGrid
End Sub
