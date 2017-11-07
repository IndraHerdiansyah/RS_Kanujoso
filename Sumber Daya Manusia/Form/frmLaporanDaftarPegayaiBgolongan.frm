VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLaporanDaftarPegawaiMenurutGolongan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Laporan Daftar Pegawai Berdasarkan Golongan"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLaporanDaftarPegayaiBgolongan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6060
   Begin VB.CommandButton cmdtutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2280
      Width           =   1905
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2280
      Width           =   1905
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
      Height          =   1035
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   5355
      Begin VB.Frame Frame3 
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   150
         Width           =   4935
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MMMM yyyy "
            Format          =   104792067
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker DTPickerAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MMMM yyyy"
            Format          =   104792067
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   4
            Top             =   315
            Visible         =   0   'False
            Width           =   255
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
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
      Left            =   3480
      Picture         =   "frmLaporanDaftarPegayaiBgolongan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2835
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLaporanDaftarPegayaiBgolongan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanDaftarPegayaiBgolongan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "frmLaporanDaftarPegawaiMenurutGolongan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    On Error GoTo errLoad
    Dim pesan As VbMsgBoxResult
    strSQL = "SELECT     TypePegawai,SUM(GOL1a) as GOL1a,SUM(GOL1b) as GOL1b,SUM(GOL1c) as GOL1c,SUM(GOL1d) as GOL1d,SUM(GOL2a) as GOL2a,SUM(GOL2b) as GOL2b,  " & _
             " sum(GOL2c) as GOL2c,SUM(GOL2d) as GOL2d,SUM(GOL3a) as GOL3a,SUM(GOL3b) as GOL3b,SUM(GOL3c) as GOL3c,SUM(GOL3d) as GOL3d,SUM(GOL4a) as GOL4a,SUM(GOL4b) as  GOL4b, " & _
             " SUM(GOL4c) as GOL4c,SUM(GOL4d) as  GOL4d " & _
             "  From dbo.v_LapGolonganPegawai " & _
             "   GROUP BY TypePegawai "
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        MsgBox "Tidak ada data  ", vbCritical, "Validasi"
        Exit Sub
    End If

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    frm_cetak_LaporanDaftarPegawaiBerDasarkanGolongan.Show
    Exit Sub
errLoad:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub DTPickerAwal_Change()
    DTPickerAwal.MaxDate = Now
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .DTPickerAwal.Value = Format(Now, "MMMM yyyy")
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub
