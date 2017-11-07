VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Pegawai"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   13560
   Begin VB.OptionButton Option2 
      Caption         =   "Tidak Aktif"
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
      Left            =   120
      TabIndex        =   43
      Top             =   3960
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Aktif"
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
      Left            =   120
      TabIndex        =   42
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   13335
      Begin VB.TextBox txtIDPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   15
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtTptLhr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10800
         MaxLength       =   50
         TabIndex        =   14
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtNIP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1680
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo dcJnsPeg 
         Height          =   330
         Left            =   1920
         TabIndex        =   17
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcPangkat 
         Height          =   330
         Left            =   3720
         TabIndex        =   18
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcGol 
         Height          =   330
         Left            =   6360
         TabIndex        =   19
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcJabatan 
         Height          =   330
         Left            =   8640
         TabIndex        =   20
         Top             =   1080
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcPddk 
         Height          =   330
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSMask.MaskEdBox meTglLahir 
         Height          =   300
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcStatusPegawai 
         Height          =   330
         Left            =   5400
         TabIndex        =   23
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dcJK 
         Height          =   330
         Left            =   9120
         TabIndex        =   24
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSMask.MaskEdBox meTglMasuk 
         Height          =   300
         Left            =   1920
         TabIndex        =   25
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcJabatanF 
         Height          =   330
         Left            =   7200
         TabIndex        =   40
         Top             =   1680
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan Fungsional"
         Height          =   210
         Index           =   13
         Left            =   7200
         TabIndex        =   41
         Top             =   1440
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl.Masuk"
         Height          =   210
         Index           =   12
         Left            =   1920
         TabIndex        =   38
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Pegawai"
         Height          =   210
         Index           =   11
         Left            =   5400
         TabIndex        =   37
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID Pegawai"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pegawai"
         Height          =   210
         Index           =   1
         Left            =   1920
         TabIndex        =   35
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
         Height          =   210
         Index           =   2
         Left            =   5280
         TabIndex        =   34
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   9120
         TabIndex        =   33
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Lahir"
         Height          =   210
         Index           =   4
         Left            =   10800
         TabIndex        =   32
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl.Lahir"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pangkat"
         Height          =   210
         Index           =   6
         Left            =   3720
         TabIndex        =   30
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Golongan"
         Height          =   210
         Index           =   7
         Left            =   6360
         TabIndex        =   29
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   210
         Index           =   8
         Left            =   8640
         TabIndex        =   28
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pendidikan Terakhir"
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
         Height          =   210
         Index           =   10
         Left            =   2520
         TabIndex        =   26
         Top             =   1440
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtParameter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   1
      Top             =   7800
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Detail"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Riwayat"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdAlamat 
      Caption         =   "&Alamat"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   8880
      TabIndex        =   6
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   10440
      TabIndex        =   7
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdtutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   12000
      TabIndex        =   8
      Top             =   7800
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   8370
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11906
            Text            =   "F1 - Cetak Data Pegawai"
            TextSave        =   "F1 - Cetak Data Pegawai"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11906
            Text            =   "F2 - Cetak SKUM"
            TextSave        =   "F2 - Cetak SKUM"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   1560
      TabIndex        =   0
      Top             =   3360
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   16
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
   Begin VB.Label lblJumData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data 0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   39
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   945
      Left            =   11760
      Picture         =   "frmDataPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataPegawai.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11775
   End
   Begin VB.Label Label2 
      Caption         =   "Cari Pegawai :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   11640
      Picture         =   "frmDataPegawai.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataPegawai.frx":459E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmDataPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim kdJnsPegawai As String, kdPangkat As String, kdGol As String, kdJabatan As String
Dim kdPendidikan As String
Dim msg As VbMsgBoxResult

Private Sub subLoadFormRiwayatPegawai()
    On Error GoTo errLoad
    mstrIdPegawai = DataGrid1.Columns(0).Value

    With frmRiwayatPegawai
        .Show
        .txtIDPegawai.Text = DataGrid1.Columns(0).Text
        .txtJenisPegawai.Text = frmDataPegawai.dcJnsPeg.Text
        .txtNamaPegawai.Text = DataGrid1.Columns(2).Text
        .txtJabatan.Text = DataGrid1.Columns(9).Text
        If DataGrid1.Columns(3) = "L" Then
            .txtSex.Text = "Laki Laki"
        Else
            .txtSex.Text = "Perempuan"
        End If
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdAlamat_Click()
    If txtIDPegawai.Text = "" Then
        MsgBox "Pilih Pegawai yang akan ditampilkan Alamatnya!", vbExclamation, "Validasi"
    Else
        frmDataAlamatPegawai.Show
    End If
End Sub

Private Sub cmdBatal_Click()
    On Error GoTo errLoad
    Call clearData
    Call subLoadGridSource
    Call subLoadDcSource
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo xxx
    If txtIDPegawai.Text = "" Then Exit Sub
    If MsgBox("Hapus data pegawai dengan nomor " & txtIDPegawai.Text & "", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    dbConn.Execute "DELETE HistoryLoginAplikasi where IdPegawai = '" & txtIDPegawai.Text & "'"
    dbConn.Execute "DELETE DataCurrentPegawai where IdPegawai = '" & txtIDPegawai.Text & "'"
    dbConn.Execute "DELETE DataPegawai where IdPegawai = '" & txtIDPegawai.Text & "'"
    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call subLoadGridSource
    Call cmdBatal_Click
    Exit Sub
xxx:
    MsgBox "Data Tidak Dapat Dihapus..", vbOKOnly, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim sStatus As String

    If Periksa("datacombo", dcJnsPeg, "Jenis pegawai kosong") = False Then Exit Sub
    If Periksa("text", txtNama, "Nama pegawai kosong") = False Then Exit Sub
    If Periksa("datacombo", dcJK, "Jenis kelamin kosong") = False Then Exit Sub
    If Periksa("datacombo", dcStatusPegawai, "Status Pegawai Kosong") = False Then Exit Sub
    If Periksa("datacombo", dcJabatan, "Jabatan Pegawai Kosong") = False Then Exit Sub
    If funcCekValidasiTgl("TglLahir", meTglLahir) <> "NoErr" Then
        MsgBox "Tanggal lahir kosong", vbExclamation, "Validasi"
        meTglLahir.SetFocus
        Exit Sub
    End If

    If sp_DataPegawai = False Then Exit Sub
    If sp_DataCurrentPegawai = False Then Exit Sub

    Call subLoadGridSource
    Call clearData
    Call cmdBatal_Click
    Call txtParameter_Change
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If txtIDPegawai.Text = "" Then
        MsgBox "Pilih Pegawai yang akan ditampilkan Detail Pegawainya!", vbExclamation, "Validasi"
        Exit Sub
    End If
    With frmDetailPegawai
        .Show
        .txtIDPegawai.Text = mstrIdPegawai
        .txtNamaLengkap.Text = txtNama.Text
    End With
End Sub

Private Sub Command2_Click()
    If txtIDPegawai.Text = "" Then
        MsgBox "Pilih Pegawai yang akan ditampilkan Riwayatnya!", vbExclamation, "Validasi"
    Else
        If dcJabatan.BoundText = "" Then MsgBox "Silahkan lengkapi jabatan pegawai", vbExclamation, "Validasi": Exit Sub
        Call subLoadFormRiwayatPegawai
        Me.Enabled = False
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    lblJumData.Caption = "Data " & DataGrid1.Bookmark & "/" & DataGrid1.ApproxCount
    With DataGrid1
        txtIDPegawai.Text = .Columns("ID Pegawai").Value
        If .Columns("KodePeg").Value = "" Then
            dcJnsPeg.BoundText = ""
        Else
            dcJnsPeg.BoundText = .Columns("KodePeg").Value
        End If
        txtNama.Text = .Columns("Nama Lengkap").Value
        dcJK.BoundText = .Columns("JK").Value
        If .Columns("Tempat Lahir").Value = "" Then
            txtTptLhr.Text = ""
        Else
            txtTptLhr.Text = .Columns("Tempat Lahir").Value
        End If
        If .Columns("Pangkat").Value = "" Then
            dcPangkat.Text = ""
        Else
            dcPangkat.Text = .Columns("Pangkat").Value
        End If
        If .Columns("Golongan").Value = "" Then
            dcGol.Text = ""
        Else
            dcGol.Text = .Columns("Golongan").Value
        End If
        If .Columns("Jabatan").Value = "" Then
            dcJabatan.Text = ""
        Else
            dcJabatan.Text = .Columns("Jabatan").Value
        End If
        If .Columns("Pendidikan").Value = "" Then
            dcPddk.Text = ""
        Else
            dcPddk.Text = .Columns("Pendidikan").Value
        End If
        If .Columns("NIP").Value = "" Then
            txtNIP.Text = ""
        Else
            txtNIP.Text = .Columns("NIP").Value
        End If
        If .Columns("Status").Value = "" Then
            dcStatusPegawai.Text = ""
        Else
            dcStatusPegawai.Text = .Columns("Status").Value
        End If
        If .Columns("Tgl. Lahir").Value = "" Then
            meTglLahir.Text = "__/__/____"
        Else
            meTglLahir.Text = .Columns("Tgl. Lahir").Value
        End If
        If .Columns("Tgl. Masuk").Value = "" Then
            meTglMasuk.Text = "__/__/____"
        Else
            meTglMasuk.Text = .Columns("Tgl. Masuk").Value
        End If

        If .Columns("Jabatan Fungsi").Value = "" Then
            dcJabatanF.Text = ""
        Else
            dcJabatanF.Text = .Columns("Jabatan Fungsi").Value
        End If
        mstrIdPegawai = txtIDPegawai.Text
    End With
    Exit Sub
End Sub

Private Sub dcGol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcGol.Text)) = 0 Then dcJabatan.SetFocus: Exit Sub
        If dcGol.MatchedWithList = True Then dcJabatan.SetFocus: Exit Sub
        strSQL = "select kdGolongan,NamaGolongan from golonganpegawai WHERE (NamaGolongan LIKE '%" & dcGol.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcGol.BoundText = rs(0).Value
        dcGol.Text = rs(1).Value
    End If
End Sub

Private Sub dcJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcJabatan.Text)) = 0 Then dcPddk.SetFocus: Exit Sub
        If dcJabatan.MatchedWithList = True Then dcPddk.SetFocus: Exit Sub
        strSQL = "select kdJabatan,Namajabatan from jabatan WHERE (Namajabatan LIKE '%" & dcJabatan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcJabatan.BoundText = rs(0).Value
        dcJabatan.Text = rs(1).Value
    End If
End Sub

Private Sub dcJabatanF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dcJK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcJK.Text)) = 0 Then txtTptLhr.SetFocus: Exit Sub
        If dcJnsPeg.MatchedWithList = True Then txtTptLhr.SetFocus: Exit Sub
        strSQL = "select Singkatan,JenisKelamin from jenisKelamin WHERE (jeniskelamin LIKE '%" & dcJK.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcJK.BoundText = rs(0).Value
        dcJK.Text = rs(1).Value
    End If
End Sub

Private Sub dcJnsPeg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcJnsPeg.Text)) = 0 Then txtNama.SetFocus: Exit Sub
        If dcJnsPeg.MatchedWithList = True Then txtNama.SetFocus: Exit Sub
        strSQL = "select kdJenisPegawai,jenispegawai from jenispegawai WHERE (jenispegawai LIKE '%" & dcJnsPeg.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcJnsPeg.BoundText = rs(0).Value
        dcJnsPeg.Text = rs(1).Value
    End If
End Sub

Private Sub dcPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcPangkat.Text)) = 0 Then dcGol.SetFocus: Exit Sub
        If dcPangkat.MatchedWithList = True Then dcGol.SetFocus: Exit Sub
        strSQL = "select kdPangkat,NamaPangkat from Pangkat WHERE (NamaPangkat LIKE '%" & dcPangkat.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPangkat.BoundText = rs(0).Value
        dcPangkat.Text = rs(1).Value
    End If
End Sub

Private Sub dcPddk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(dcPddk.Text)) = 0 Then txtNIP.SetFocus: Exit Sub
        If dcPddk.MatchedWithList = True Then txtNIP.SetFocus: Exit Sub
        strSQL = "select kdPendidikan,pendidikan from pendidikan WHERE (pendidikan LIKE '%" & dcPddk.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPddk.BoundText = rs(0).Value
        dcPddk.Text = rs(1).Value
    End If
End Sub

Private Sub dcStatusPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJabatanF.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    centerForm Me, MDIUtama
    Option1.Value = True
    Call cmdBatal_Click
End Sub

Sub clearData()
    On Error Resume Next
    txtIDPegawai.Text = ""
    txtNama.Text = ""
    txtNIP.Text = ""
    txtTptLhr.Text = ""
    dcJK.BoundText = ""
    dcGol.BoundText = ""
    dcJabatan.BoundText = ""
    dcJnsPeg.BoundText = ""
    dcPangkat.BoundText = ""
    dcPddk.BoundText = ""
    dcStatusPegawai.BoundText = ""
    meTglLahir.Text = "__/__/____"
    meTglMasuk.Text = "__/__/____"
    dcJabatanF.BoundText = ""
End Sub

Private Sub subLoadDcSource()
    strSQL = "select kdJenisPegawai,jenispegawai from jenispegawai order by jenispegawai "
    Call msubDcSource(dcJnsPeg, rs, strSQL)

    strSQL = "select kdGolongan,NamaGolongan from golonganpegawai order by namagolongan"
    Call msubDcSource(dcGol, rs, strSQL)

    strSQL = "select kdPangkat,NamaPangkat from Pangkat order by namapangkat"
    Call msubDcSource(dcPangkat, rs, strSQL)

    strSQL = "select kdJabatan,Namajabatan from jabatan order by namajabatan"
    Call msubDcSource(dcJabatan, rs, strSQL)

    strSQL = "select kdPendidikan,pendidikan from pendidikan order by pendidikan"
    Call msubDcSource(dcPddk, rs, strSQL)

    strSQL = "select kdStatus,Status from StatusPegawai order by Status"
    Call msubDcSource(dcStatusPegawai, rs, strSQL)

    strSQL = "select Singkatan,JenisKelamin from JenisKelamin order by JenisKelamin"
    Call msubDcSource(dcJK, rs, strSQL)

    strSQL = "select KdDetailJabatanF,DetailjabatanF from DetailJabatanFungsional order by DetailJabatanF"
    Call msubDcSource(dcJabatanF, rs, strSQL)
End Sub



Private Sub meTglLahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meTglMasuk.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        FrmCetakDataPegawai.Show
    End If
    If KeyCode = vbKeyF2 Then
        mstrIdPegawai = DataGrid1.Columns(0).Value
        strSQL = "Select * from V_SKUM where IdPegawai = '" & mstrIdPegawai & "' "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            MsgBox "Silahkan lengkapi seluruh data pegawai " & DataGrid1.Columns("Nama Lengkap").Value & " ", vbCritical, "Validasi"
            Exit Sub
        Else
            frmCetakSKUM.Show

        End If
    End If
End Sub

Private Sub meTglLahir_LostFocus()
    On Error GoTo errTglLahir
    If meTglLahir.Text = "__/__/____" Then Exit Sub
    If funcCekValidasiTgl("TglLahir", meTglLahir) <> "NoErr" Then Exit Sub
    Exit Sub
errTglLahir:
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
    & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Sub meTglMasuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPangkat.SetFocus
End Sub

Private Sub meTglMasuk_LostFocus()
    On Error GoTo errTglLahir
    If meTglMasuk.Text = "__/__/____" Then Exit Sub
    If funcCekValidasiTgl("TglLahir", meTglMasuk) <> "NoErr" Then Exit Sub
    Exit Sub
errTglLahir:
    MsgBox "Sudahkah anda mengganti Regional Setting komputer anda menjadi 'Indonesia'?" _
    & vbNewLine & "Bila sudah hubungi Administrator dan laporkan pesan kesalahan berikut:" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Sub Option1_Click()
    Call subLoadGridSource
End Sub

Private Sub Option2_Click()
    Call subLoadGridSource
End Sub

Private Sub txtIDPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJK.SetFocus
End Sub

Private Sub txtNIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcStatusPegawai.SetFocus
End Sub

Private Sub txtParameter_Change()
    Call subLoadGridSource
End Sub

Private Sub txtTptLhr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then meTglLahir.SetFocus
End Sub

Sub subLoadGridSource()
    On Error GoTo errLoad
    If Option1.Value = True Then
        Set rs = Nothing
        strSQL = "select * from V_M_DataPegawai WHERE KdStatus='01' AND [Nama Lengkap] LIKE '%" & txtParameter.Text & "%' order by [Nama Lengkap]"
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        Set DataGrid1.DataSource = rs
        lblJumData.Caption = "Data " & DataGrid1.Bookmark & "/" & DataGrid1.ApproxCount
        Call SetDataGrid
    ElseIf Option2.Value = True Then
        Set rs = Nothing
        strSQL = "select * from V_M_DataPegawai WHERE KdStatus<>'01' AND [Nama Lengkap] LIKE '%" & txtParameter.Text & "%' order by [Nama Lengkap]"
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        Set DataGrid1.DataSource = rs
        lblJumData.Caption = "Data " & DataGrid1.Bookmark & "/" & DataGrid1.ApproxCount
        Call SetDataGrid
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetDataGrid()
    With DataGrid1
        .Columns(0).Width = 1200
        .Columns(1).Width = 0
        .Columns(2).Width = 2500
        .Columns(3).Width = 250
        .Columns(4).Width = 1200
        .Columns(5).Width = 1200
        .Columns(6).Width = 0
        .Columns(7).Width = 2000
        .Columns(8).Width = 1000
        .Columns(9).Width = 2000
        .Columns(10).Width = 1000
        .Columns(11).Width = 1000
        .Columns(12).Width = 1500
        .Columns(13).Width = 2000
        .Columns(14).Width = 0
        .Columns(15).Width = 0
    End With
End Sub

Private Sub txtTptLhr_LostFocus()
    txtTptLhr.Text = StrConv(txtTptLhr.Text, vbProperCase)
End Sub

Private Function sp_DataPegawai() As Boolean
    sp_DataPegawai = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIDPegawai.Text)
        .Parameters.Append .CreateParameter("KdJenisPegawai", adChar, adParamInput, 3, IIf(dcJnsPeg.BoundText = "", Null, dcJnsPeg.BoundText))
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, txtNama.Text)
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, dcJK.BoundText)
        .Parameters.Append .CreateParameter("TempatLahir", adVarChar, adParamInput, 50, IIf(Len(Trim(txtTptLhr.Text)) = 0, Null, Trim(txtTptLhr.Text)))
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(meTglLahir, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , IIf(meTglMasuk = "__/__/____", Null, Format(meTglMasuk, "yyyy/MM/dd")))
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
            If Not IsNull(.Parameters("OutKode").Value) Then txtIDPegawai = .Parameters("OutKode").Value
            mstrIdPegawai = txtIDPegawai.Text
            MsgBox "Penyimpanan Data Berhasil..", vbInformation, "Informasi"
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
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIDPegawai.Text)
        .Parameters.Append .CreateParameter("KdPangkat", adVarChar, adParamInput, 2, IIf(dcPangkat.BoundText = "", Null, dcPangkat.BoundText))
        .Parameters.Append .CreateParameter("KdGolongan", adVarChar, adParamInput, 2, IIf(dcGol.BoundText = "", Null, dcGol.BoundText))
        .Parameters.Append .CreateParameter("KdPendidikan", adChar, adParamInput, 2, IIf(dcPddk.BoundText = "", Null, dcPddk.BoundText))
        .Parameters.Append .CreateParameter("KdJabatan", adVarChar, adParamInput, 5, IIf(dcJabatan.BoundText = "", Null, dcJabatan.BoundText))
        .Parameters.Append .CreateParameter("NIP", adVarChar, adParamInput, 30, IIf(txtNIP.Text = "", Null, txtNIP.Text))
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, dcStatusPegawai.BoundText)
        .Parameters.Append .CreateParameter("KdDetailJabatanF", adChar, adParamInput, 5, IIf(dcJabatanF.BoundText = "", Null, dcJabatanF.BoundText))
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
