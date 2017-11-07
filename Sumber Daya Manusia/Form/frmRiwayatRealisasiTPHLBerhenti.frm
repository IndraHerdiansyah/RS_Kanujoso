VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRiwayatRealisasiTPHLBerhenti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Realisasi Pemberhentian Tenaga Pegawai Harian Lepas (TPHL)"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   Icon            =   "frmRiwayatRealisasiTPHLBerhenti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   9735
   Begin VB.Frame Frame2 
      Caption         =   "Pegawai"
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
      TabIndex        =   22
      Top             =   1080
      Width           =   9495
      Begin VB.TextBox txtPendidikan 
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
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   29
         Top             =   600
         Width           =   2295
      End
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
         TabIndex        =   24
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtTempatlahir 
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   23
         Top             =   600
         Width           =   2295
      End
      Begin MSMask.MaskEdBox meTglLahir 
         Height          =   300
         Left            =   5760
         TabIndex        =   27
         Top             =   600
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pendidikan"
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
         Left            =   7080
         TabIndex        =   30
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
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
         Left            =   5760
         TabIndex        =   28
         Top             =   360
         Width           =   960
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
         TabIndex        =   26
         Top             =   360
         Width           =   1050
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
         Left            =   3360
         TabIndex        =   25
         Top             =   360
         Width           =   930
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   11
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
      TabIndex        =   8
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Realisasi Pemberhentian"
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
      TabIndex        =   13
      Top             =   2160
      Width           =   9495
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
         Left            =   240
         MaxLength       =   100
         TabIndex        =   6
         Top             =   2160
         Width           =   3855
      End
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
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtTugasKerja 
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
         Left            =   4200
         MaxLength       =   150
         TabIndex        =   7
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox txtNoSK 
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
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1320
         Width           =   3495
      End
      Begin MSDataListLib.DataCombo dcStatus 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
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
      Begin MSComCtl2.DTPicker dtpTglSK 
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   1320
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
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy"
         Format          =   67960832
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpSetuju 
         Height          =   330
         Left            =   2640
         TabIndex        =   1
         Top             =   600
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
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy"
         Format          =   67960832
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpBerlaku 
         Height          =   330
         Left            =   5040
         TabIndex        =   2
         Top             =   600
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
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy"
         Format          =   67960832
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Berlaku"
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
         Left            =   5040
         TabIndex        =   33
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Disetujui"
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
         Left            =   2640
         TabIndex        =   32
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   31
         Top             =   1920
         Width           =   840
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
         Left            =   6240
         TabIndex        =   19
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
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tugas Pekerjaan"
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
         Left            =   4200
         TabIndex        =   17
         Top             =   1920
         Width           =   1200
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
         Left            =   2640
         TabIndex        =   15
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Status Disetujui"
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
         Width           =   1125
      End
   End
   Begin MSDataGridLib.DataGrid dgData 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
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
      Picture         =   "frmRiwayatRealisasiTPHLBerhenti.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatRealisasiTPHLBerhenti.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatRealisasiTPHLBerhenti.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatRealisasiTPHLBerhenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''add by splakuk 2010/8/14
Option Explicit

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadRiwayatRealisasi
    dcStatus.SetFocus
End Sub

Private Sub cmdHapus_Click()
On Error GoTo errHapus
    If txtIdPegawai.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatRealisasiUsulan WHERE NoRiwayat='" & txtNoRiwayat.Text & "' and IdPegawai = '" & txtIdPegawai.Text & "' "
    dbConn.Execute strSQL
    If sp_Riwayat("D") = False Then Exit Sub
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    
    Call cmdBatal_Click
Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    If mstrIdPegawai = "" Then Exit Sub
    If Periksa("datacombo", dcStatus, "Silahkan isi status realisasi ") = False Then Exit Sub
    If sp_Riwayat("A") = False Then Exit Sub
    If sp_RiwayatRealisasi = False Then Exit Sub

    
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Call cmdBatal_Click

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
If dgData.ApproxCount = 0 Then Exit Sub
    With dgData
        dcStatus.BoundText = .Columns(6).Value
        If IsNull(.Columns(2).Value) Then dtpTglSK.Value = Null Else dtpTglSK.Value = .Columns(2).Value
        If IsNull(.Columns(3).Value) Then txtNoSK.Text = "" Else txtNoSK.Text = .Columns(3).Value
        If IsNull(.Columns(4).Value) Then txtTTD.Text = "" Else txtTTD.Text = .Columns(4).Value
        If IsNull(.Columns(5).Value) Then txtTugasKerja.Text = "" Else txtTugasKerja.Text = .Columns(5).Value
        If IsNull(.Columns(10).Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = .Columns(10).Value
        
        txtNoRiwayat.Text = .Columns(0).Value
        txtIdPegawai.Text = .Columns(1).Value
    End With
End Sub

Private Sub dcStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpSetuju.SetFocus
End Sub

Private Sub dtpBerlaku_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglSK.SetFocus
End Sub

Private Sub dtpSetuju_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpBerlaku.SetFocus
End Sub

Private Sub dtpTglSK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNoSK.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call SetComboStatus
    Call SetComboAlasanStatus
    Call subLoadRiwayatRealisasi
End Sub

Private Sub subLoadRiwayatRealisasi()
On Error GoTo errload
    strSQL = "SELECT dbo.RiwayatUsulan.NoRiwayat, dbo.RiwayatUsulan.IdPegawai, dbo.RiwayatUsulan.TglSK, dbo.RiwayatUsulan.NoSK, " & _
             "dbo.RiwayatUsulan.TandaTanganSK, dbo.RiwayatUsulan.TugasPekerjaan, dbo.RiwayatUsulan.KdStatusUsulan, dbo.StatusPegawai.Status, " & _
             "dbo.RiwayatUsulan.KdAlasanStatus, dbo.AlasanStatusPegawai.AlasanStatus, dbo.RiwayatUsulan.KeteranganLainnya, " & _
             "dbo.RiwayatUsulan.NoRiwayatRealisasi " & _
             "FROM dbo.RiwayatUsulan INNER JOIN " & _
             "dbo.StatusPegawai ON dbo.RiwayatUsulan.KdStatusUsulan = dbo.StatusPegawai.KdStatus INNER JOIN " & _
             "dbo.AlasanStatusPegawai ON dbo.RiwayatUsulan.KdAlasanStatus = dbo.AlasanStatusPegawai.KdAlasanStatus " & _
             "WHERE (dbo.RiwayatUsulan.KdAlasanStatus IS NOT NULL) " & _
             "WHERE dbo.RiwayatUsulan.IdPegawai = '" & mstrIdPegawai & "' "
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgData.DataSource = rsb
    With dgData

        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i

        .Columns(0).Width = 1000 'noriwayat
        .Columns(2).Width = 1000 'tglsk
        .Columns(3).Width = 1500 'nosk
        .Columns(4).Width = 1500 'ttdsk
        .Columns(5).Width = 2000 'tugaskerja
        .Columns(7).Width = 1000 'status
        .Columns(10).Width = 2000 'keterangan
        .Columns(9).Width = 1500 'alasan status
    
    End With
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub subClearData()
On Error Resume Next
    txtIdPegawai.Text = ""
    txtNoRiwayat.Text = ""
    dcStatus.BoundText = ""
    dtpTglSK.Value = Format(Now, "dd/mmmm/yyyy")
    txtTTD.Text = ""
    txtTugasKerja.Text = ""
    txtNoSK.Text = ""
    txtKeterangan.Text = ""
End Sub

Private Sub txtNoSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtTTD.SetFocus
End Sub

 Sub SetComboStatus()
 On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from StatusPegawai order by Status ", dbConn, , adLockOptimistic
    Set dcStatus.RowSource = rs
    dcStatus.ListField = rs.Fields(1).Name
    dcStatus.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTugasKerja.SetFocus
End Sub

Private Sub txtTTD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Function sp_Riwayat(f_Status) As Boolean
On Error GoTo hell
    sp_Riwayat = True
    Set dbcmd = New ADODB.Command
    With dbcmd
    
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        If txtNoRiwayat = "" Then
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtNoRiwayat.Text)
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
            If Not IsNull(.Parameters("Status").Value) Then txtNoRiwayat.Text = .Parameters("OutputNoRiwayat").Value
        End If
        
        
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_RiwayatRealisasi() As Boolean
On Error GoTo hell
    sp_RiwayatRealisasi = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtNoRiwayat.Text))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        If IsNull(dtpTglSK.Value) Then
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Format(dtpTglSK.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("NoSK", adVarChar, adParamInput, 30, IIf(txtNoSK.Text = "", Null, Trim(txtNoSK.Text)))
        .Parameters.Append .CreateParameter("TandaTanganSK", adVarChar, adParamInput, 50, IIf(txtTTD.Text = "", Null, Trim(txtTTD.Text)))
        .Parameters.Append .CreateParameter("TugasPekerjaan", adVarChar, adParamInput, 150, IIf(txtTugasKerja.Text = "", Null, Trim(txtTugasKerja.Text)))
        .Parameters.Append .CreateParameter("KdStatusUsulan", adChar, adParamInput, 2, dcStatus.BoundText)
        .Parameters.Append .CreateParameter("KdAlasanStatus", adTinyInt, adParamInput, , dcAlasan.BoundText)
        .Parameters.Append .CreateParameter("KdDKategoryPUsulan", adVarChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("GajiPokokUsulan", adCurrency, adParamInput, , Null)
        .Parameters.Append .CreateParameter("KdPangkatUsulan", adVarChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("NoRiwayatRealisasi", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TotalPaguUsulan", adCurrency, adParamInput, , Null)
        .Parameters.Append .CreateParameter("BankPenyalur", adVarChar, adParamInput, 50, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RiwayatUsulan"
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

Private Sub txtTugasKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
