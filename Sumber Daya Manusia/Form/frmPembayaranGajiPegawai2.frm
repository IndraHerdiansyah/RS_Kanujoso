VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPembayaranGajiPegawai2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pembayaran Gaji Pegawai"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10725
   Icon            =   "frmPembayaranGajiPegawai2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10725
   Begin VB.TextBox txtJumlah 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   33
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   10815
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   10455
         Begin VB.TextBox txtMasaKErja 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9600
            TabIndex        =   37
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtPEndidikan 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtTypePegawai 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8880
            TabIndex        =   35
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtGolongan 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8880
            TabIndex        =   34
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtJabatan 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5880
            TabIndex        =   31
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txtKeterangan 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2280
            TabIndex        =   9
            Top             =   960
            Width           =   4215
         End
         Begin VB.TextBox txtPegawai 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6600
            TabIndex        =   8
            Top             =   960
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dtpTglBayar 
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   114032643
            UpDown          =   -1  'True
            CurrentDate     =   37760
         End
         Begin MSDataListLib.DataCombo dcNamaPegawai 
            Height          =   330
            Left            =   2280
            TabIndex        =   24
            Top             =   360
            Width           =   3495
            _ExtentX        =   6165
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
         Begin VB.Label Label8 
            Caption         =   "Jabatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   32
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Keterangan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   11
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Nama Pegawai"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   10
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Tgl Pembayaran"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Potongan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   5400
         TabIndex        =   26
         Top             =   1560
         Width           =   5175
         Begin MSFlexGridLib.MSFlexGrid dgPotongan 
            Height          =   2175
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3836
            _Version        =   393216
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
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pendapatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   5175
         Begin MSFlexGridLib.MSFlexGrid dgPendapatan 
            Height          =   2175
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3836
            _Version        =   393216
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
      End
      Begin VB.Frame fraPegawai 
         Caption         =   "Pegawai"
         Height          =   2175
         Left            =   10560
         TabIndex        =   22
         Top             =   7680
         Visible         =   0   'False
         Width           =   6255
         Begin MSDataGridLib.DataGrid dgPegawai 
            Height          =   1815
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   3201
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
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   10455
         Begin VB.TextBox txtTotalBersih 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            TabIndex        =   18
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtTotalPotongan 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   17
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtTotalPendapatan 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Total Pendapatan Bersih"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5400
            TabIndex        =   21
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Total Potongan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   20
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Total Pendapatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
      End
      Begin TabDlg.SSTab SSTabPembayaranGaji 
         Height          =   4335
         Left            =   11520
         TabIndex        =   6
         Top             =   2040
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7646
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Pendapatan"
         TabPicture(0)   =   "frmPembayaranGajiPegawai2.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Potongan"
         TabPicture(1)   =   "frmPembayaranGajiPegawai2.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgPotongan1"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid dgPotongan1 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   7
            Top             =   360
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   6800
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
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   4920
         Width           =   10455
         Begin VB.CommandButton cmdSimpan 
            Caption         =   "Simpan"
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
            Left            =   7440
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdBatal 
            Caption         =   "Batal"
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
            Left            =   6120
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdCetak 
            Caption         =   "Cetak Struk"
            Enabled         =   0   'False
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
            Left            =   4800
            TabIndex        =   30
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdTutup 
            Caption         =   "Tutup"
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
            Left            =   9000
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "F1 Cetak Struk / Slip Gaji"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
            Width           =   3135
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
      Left            =   8880
      Picture         =   "frmPembayaranGajiPegawai2.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPembayaranGajiPegawai2.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmPembayaranGajiPegawai2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
    dcNamaPegawai.BoundText = ""
    cmdCetak.Enabled = False
    cmdSimpan.Enabled = True
    dcNamaPegawai.Enabled = True
End Sub

Private Sub cmdCetak_Click()
    frmCetakStrukGajiPegawai.Show
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If Periksa("text", dcNamaPegawai, "Silahkan isi Nama Pegawai") = False Then Exit Sub

    With dgPendapatan
        strSQL = "delete from PembayaranGajiPegawai where idpegawai='" & dcNamaPegawai.BoundText & "' and TglPembayaran='" & Format(dtpTglBayar, "yyyy/MM/dd") & "'"
        Call msubRecFO(rs, strSQL)
        For i = 1 To .Rows - 1
            If .Rows - 1 = 0 Then MsgBox "Lengkapi data riwayat gaji ", vbExclamation, "Validasi": Exit Sub
            If sp_SimpanPembayaranGajiPegawai(.TextMatrix(i, 0), .TextMatrix(i, 1), .TextMatrix(i, 3), _
                .TextMatrix(i, 4), "A") = False Then Exit Sub
            Next i
        End With

        With dgPotongan
            strSQL = "delete from PembayaranPotonganGajiPegawai where idpegawai='" & dcNamaPegawai.BoundText & "' and TglPembayaran='" & Format(dtpTglBayar, "yyyy/MM/dd") & "'"
            Call msubRecFO(rs, strSQL)
            For i = 1 To .Rows - 1
                If .Rows - 1 = 0 Then MsgBox "Lengkapi data komponen potongan ", vbExclamation, "Validasi": Exit Sub
                If sp_SimpanPembayaranPotonganGaji(.TextMatrix(i, 0), .TextMatrix(i, 1), .TextMatrix(i, 3), _
                    .TextMatrix(i, 4), "A") = False Then Exit Sub
                Next i
            End With

            strSQL = "Select * from V_4insertPembayaranGajidNPotongan where IdPegawai='" & dcNamaPegawai.BoundText & "' and TglPembayaran='" & Format(dtpTglBayar, "yyyy/MM/dd") & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                For i = 1 To rs.RecordCount
                    If sp_SimpanPembayaranDanPotonganGaji(rs.Fields("IdPegawai"), rs.Fields("KdKomponenGaji"), rs.Fields("KomponenGaji"), rs.Fields("Jumlah"), rs.Fields("KdKomponenPotonganGaji"), _
                        rs.Fields("KomponenPotonganGaji"), rs.Fields("JumlahPotongan"), "A") = False Then Exit Sub
                    Next i

                End If

'                strSQL = "select*from V_PPH21 where tglPembayaran = '" & Format(dtpTglBayar.Value, "yyyy/MM/dd") & "' and idPegawai = '" & dcNamaPegawai.BoundText & "'"
'                Call msubRecFO(rs, strSQL)
'                If rs.EOF = False Then
'                    If sp_SimpanPotonganPPH21 = False Then Exit Sub
'                Else
'                    MsgBox "Data potongan PPH belum dilengkapi", vbInformation, "Informasi"
'                End If
                MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
                dcNamaPegawai.Enabled = False
                cmdSimpan.Enabled = False
                cmdCetak.Enabled = True
                Exit Sub
hell:
                Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
'    cmdSimpan.Enabled = True
    Unload Me
End Sub

Private Sub dcNamaPegawai_Change()
    strSQL = "select * from V_PegawaiUntukPenggajian where idpegawai ='" & dcNamaPegawai.BoundText & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount > 0 Then
        txtTypePegawai.Text = IIf(IsNull(rs!KdTypePegawai), "", rs!KdTypePegawai)
        txtGolongan.Text = IIf(IsNull(rs!KdGolongan), "", rs!KdGolongan) 'rs!KdGolongan
        txtJabatan.Tag = IIf(IsNull(rs!kdJabatan), "", rs!kdJabatan) 'rs!kdJabatan
        txtJabatan.Text = IIf(IsNull(rs!NamaJabatan), "", rs!NamaJabatan) 'rs!kdJabatan
        txtPEndidikan.Text = IIf(IsNull(rs!KdKualifikasiJurusan), "", rs!KdKualifikasiJurusan) 'rs!KdKualifikasiJurusan
        txtMasaKErja.Text = IIf(IsNull(rs!MasaKerja), "", rs!MasaKerja) 'rs!MasaKerja
        dcNamaPegawai.Text = IIf(IsNull(rs!NamaLengkap), "", rs!NamaLengkap) 'rs!MasaKerja
        txtPegawai.Text = dcNamaPegawai.BoundText
    End If
    Call subLoadGridSource
    Call hitungPendapatan
End Sub

Private Sub dcNamaPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        strSQL = "SELECT IdPegawai, NamaLengkap FROM DataPegawai where NamaLengkap like '%" & dcNamaPegawai.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount > 0 Then
            dcNamaPegawai.BoundText = rs(0)
            
        End If
    End If
End Sub

Private Sub dgPendapatan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    'If KeyAscii = 13 Then
        If dgPendapatan.Col = 3 Then
            'Debug.Print KeyAscii
            txtJumlah.Visible = True
            txtJumlah.SetFocus
            txtJumlah.Tag = "Pendapatan"
            txtJumlah.Move Frame1.Left + Frame5.Left + dgPendapatan.Left + (dgPendapatan.ColWidth(2) * 1), Frame1.Top + Frame5.Top + dgPendapatan.Top + (dgPendapatan.RowHeight(1) * dgPendapatan.row), dgPendapatan.ColWidth(2), dgPendapatan.RowHeight(1)
            
            
            If CDbl(dgPendapatan.TextMatrix(dgPendapatan.row, 3)) > 0 Then
                txtJumlah.Text = dgPendapatan.TextMatrix(dgPendapatan.row, 3)
                txtJumlah.SelStart = 0
                txtJumlah.SelLength = Len(txtJumlah.Text)
            Else
                txtJumlah.Text = Chr(KeyAscii)
                txtJumlah.SelStart = Len(txtJumlah.Text)
            End If
        End If
    'End If
    Exit Sub
hell:
    dgPendapatan.TextMatrix(dgPendapatan.row, 3) = "0"
End Sub

Private Sub dgPotongan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If dgPotongan.Col = 3 Then
'        Debug.Print KeyAscii
        If txtJumlah.Text = "" Then txtJumlah.Text = "0"
        txtJumlah.Visible = True
        txtJumlah.SetFocus
        txtJumlah.Tag = "Potongan"
        txtJumlah.Move Frame1.Left + Frame6.Left + dgPotongan.Left + (dgPotongan.ColWidth(2) * 1), Frame1.Top + Frame6.Top + dgPotongan.Top + (dgPotongan.RowHeight(1) * dgPotongan.row), dgPotongan.ColWidth(2), dgPotongan.RowHeight(1)
        
        If CDbl(dgPotongan.TextMatrix(dgPotongan.row, 3)) > 0 Then
            txtJumlah.Text = dgPotongan.TextMatrix(dgPotongan.row, 3)
            txtJumlah.SelStart = 0
            txtJumlah.SelLength = Len(txtJumlah.Text)
        Else
            txtJumlah.Text = Chr(KeyAscii)
            txtJumlah.SelStart = Len(txtJumlah.Text)
        End If
    End If
    Exit Sub
hell:
    dgPotongan.TextMatrix(dgPotongan.row, 3) = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglBayar.Value = Format(Now, "yyyy/MMMM/dd")
    Call subLoadDcSource
    'Call subLoadMaxDate
    Call subLoadGridSource

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtJumlah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtJumlah.Text = "" Then txtJumlah.Text = "0"
        txtJumlah.Visible = False
        If txtJumlah.Tag = "Pendapatan" Then
            dgPendapatan.TextMatrix(dgPendapatan.row, 3) = txtJumlah.Text
            If dgPendapatan.Visible = True Then dgPendapatan.SetFocus
        End If
        If txtJumlah.Tag = "Potongan" Then
            dgPotongan.TextMatrix(dgPotongan.row, 3) = txtJumlah.Text
            If dgPotongan.Visible = True Then dgPotongan.SetFocus
        End If
        Call hitungPendapatan
        cmdSimpan.Enabled = True
    End If
End Sub

Private Sub txtPegawai_Change()
    fraPegawai.Visible = True
    Call subLoadGridSource
    Call hitungPendapatan
End Sub

Public Sub subLoadDcSource()
    On Error GoTo errLoad
    Call msubDcSource(dcNamaPegawai, rs, "SELECT IdPegawai, NamaLengkap FROM DataPegawai ORDER BY NamaLengkap")
    If rs.EOF = False Then dcNamaPegawai.BoundText = rs(0).Value

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Sub subLoadGridSource()
    On Error GoTo bawah
    Dim subStrKdKomponen() As String

'    txtJabatan.Text = ""
    Set rs = Nothing
    'strSQL = "SELECT idPegawai, KdKomponenGaji, KomponenGaji, Jumlah, Keterangan FROM V_RiwayatGaji where idpegawai='" & dcNamaPegawai.BoundText & "'"
'    strSQL = "select * from V_PendapatanGajiPegawai where idpegawai='" & dcNamaPegawai.BoundText & "'"
    strSQL = "select * from v_penggajianPegawai where idpegawai='" & dcNamaPegawai.BoundText & "' and KomponenGaji is not null and Jenis='Pendapatan'"
    Call msubRecFO(rs, strSQL)
    Call subSetGridPendapatan
    subBolTampil = True
    ReDim Preserve subStrKdKomponen(rs.RecordCount)
    subIntJmlKomponen = 1
    dgPendapatan.Rows = rs.RecordCount + 2
    For i = 1 To rs.RecordCount
        With dgPendapatan
            .TextMatrix(i, 0) = rs("idPegawai")
            .TextMatrix(i, 1) = rs("KdKomponenGaji")
            .TextMatrix(i, 2) = rs("KomponenGaji")
            .TextMatrix(i, 3) = rs("Jumlah")
        End With
        rs.MoveNext
    Next i

    Set rs = Nothing
    'strSQL = "SELECT idPegawai, KdKomponenPotonganGaji, KomponenPotonganGaji, JumlahPotongan, Keterangan FROM V_PotonganGaji where idpegawai='" & dcNamaPegawai.BoundText & "'"
'    strSQL = "select * from V_PotonganGajiPegawai where idpegawai ='" & dcNamaPegawai.BoundText & "'"
    strSQL = "select * from v_penggajianPegawai where idpegawai='" & dcNamaPegawai.BoundText & "' and KomponenGaji is not null and Jenis='Potongan'"
    Call msubRecFO(rs, strSQL)
    Call subSetGridPotongan
    subBolTampil = True
    ReDim Preserve subStrKdKomponen(rs.RecordCount)
    subIntJmlKomponen = 1
    dgPotongan.Rows = rs.RecordCount + 2
    For i = 1 To rs.RecordCount
        With dgPotongan
            .TextMatrix(i, 0) = rs("idPegawai")
            .TextMatrix(i, 1) = rs("KdKomponenGaji")
            .TextMatrix(i, 2) = rs("KomponenGaji")
            .TextMatrix(i, 3) = rs("Jumlah")
'            txtJabatan.Text = rs!NamaJabatan
        End With
        rs.MoveNext
    Next i
    
    Exit Sub
bawah:
End Sub

Private Sub subSetGridPendapatan()
    On Error Resume Next
    With dgPendapatan
        .clear
        .Cols = 5
        .Rows = 1
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 0
        
        .TextMatrix(0, 2) = "Komponen Gaji"
        .TextMatrix(0, 3) = "Jumlah"
        .TextMatrix(0, 4) = "Keterangan"
    End With
End Sub

Private Sub subSetGridPotongan()
    On Error Resume Next
    With dgPotongan
        .clear
        .Cols = 5
        .Rows = 1
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 0
        
        .TextMatrix(0, 2) = "Komponen Potongan Gaji"
        .TextMatrix(0, 3) = "Jumlah Potongan"
        .TextMatrix(0, 4) = "Keterangan"
    End With
End Sub

Public Sub hitungPendapatan()
On Error GoTo hell
    If dgPendapatan.Rows = 1 Then Exit Sub
    If dgPotongan.Rows = 1 Then Exit Sub
    Dim i As Integer
    txtTotalPendapatan.Text = 0
    For i = 1 To dgPendapatan.Rows - 1
        If dgPendapatan.TextMatrix(i, 3) = "" Then dgPendapatan.TextMatrix(i, 3) = "0"
        txtTotalPendapatan.Text = txtTotalPendapatan.Text + CDbl(dgPendapatan.TextMatrix(i, 3))
    Next i
    txtTotalPotongan.Text = 0
    For i = 1 To dgPotongan.Rows - 1
        If dgPotongan.TextMatrix(i, 3) = "" Then dgPotongan.TextMatrix(i, 3) = "0"
        txtTotalPotongan.Text = txtTotalPotongan.Text + CDbl(dgPotongan.TextMatrix(i, 3))
    Next i
    txtTotalBersih.Text = txtTotalPendapatan.Text - txtTotalPotongan.Text
    Exit Sub
hell:
    txtTotalPendapatan.Text = "Error"
    txtTotalPotongan.Text = "Error"
    txtTotalBersih.Text = "Error"
End Sub

Private Function sp_SimpanPembayaranGajiPegawai(f_IdPegawai As String, f_KdKomponenGaji As String, f_Jml As Currency, f_Keterangan As String, f_status As String) As Boolean
    On Error GoTo hell
    sp_SimpanPembayaranGajiPegawai = True
    
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglPembayaran", adDate, adParamInput, , Format(dtpTglBayar.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_IdPegawai)
        .Parameters.Append .CreateParameter("KdKomponenGaji", adChar, adParamInput, 2, f_KdKomponenGaji)
        .Parameters.Append .CreateParameter("Jumlah", adCurrency, adParamInput, , f_Jml)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(f_Keterangan = "", Null, f_Keterangan))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("kdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_PembayaranGajiPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam proses pemasukan data", vbCritical, "Validasi"
            sp_SimpanPembayaranGajiPegawai = False
        Else
        
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_SimpanPembayaranPotonganGaji(f_IdPegawai As String, f_KdKomponenPot As String, f_Jml As Currency, f_Keterangan As String, f_status As String) As Boolean
    On Error GoTo hell
    sp_SimpanPembayaranPotonganGaji = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglPembayaran", adDate, adParamInput, , Format(dtpTglBayar, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_IdPegawai)
        .Parameters.Append .CreateParameter("KdKomponenPotonganGaji", adChar, adParamInput, 2, f_KdKomponenPot)
        .Parameters.Append .CreateParameter("Jumlah", adCurrency, adParamInput, , f_Jml)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 30, IIf(f_Keterangan = "", Null, f_Keterangan))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("kdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_PembayaranPotonganGajiPegawai"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam proses pemasukan data", vbCritical, "Validasi"
            sp_SimpanPembayaranPotonganGaji = False
        Else
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_SimpanPembayaranDanPotonganGaji(f_IdPegawai As String, f_KdKompGaji As String, f_KompGaji As String, f_Jml As Currency, _
    f_KdKompPot As String, f_KompPot As String, f_JmlPot As Currency, f_status As String) As Boolean
    On Error GoTo hell
    Dim j As Integer
    sp_SimpanPembayaranDanPotonganGaji = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglPembayaran", adDate, adParamInput, , Format(dtpTglBayar, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_IdPegawai)
        .Parameters.Append .CreateParameter("KdKomponenGaji", adChar, adParamInput, 2, f_KdKompGaji)
        .Parameters.Append .CreateParameter("KomponenGaji", adVarChar, adParamInput, 50, f_KompGaji)
        .Parameters.Append .CreateParameter("Jumlah", adCurrency, adParamInput, , f_Jml)
        .Parameters.Append .CreateParameter("KdKomponenPotonganGaji", adChar, adParamInput, 2, f_KdKompPot)
        .Parameters.Append .CreateParameter("KomponenPotonganGaji", adVarChar, adParamInput, 50, f_KompPot)
        .Parameters.Append .CreateParameter("JumlahPotongan", adCurrency, adParamInput, , f_JmlPot)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 30, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("kdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_PembayaranGajiDanPotonganPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam proses pemasukan data", vbCritical, "Validasi"
            sp_SimpanPembayaranDanPotonganGaji = False
        Else
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_SimpanPotonganPPH21() As Boolean
    On Error GoTo hell
    sp_SimpanPotonganPPH21 = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglPembayaran", adDate, adParamInput, , Format(dtpTglBayar, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, dcNamaPegawai.BoundText)

        .ActiveConnection = dbConn
        .CommandText = "AUD_PemotonganPajakPPh21"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam proses pemasukan data", vbCritical, "Validasi"
            sp_SimpanPotonganPPH21 = False
        Else
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
hell:
    Call msubPesanError
End Function
