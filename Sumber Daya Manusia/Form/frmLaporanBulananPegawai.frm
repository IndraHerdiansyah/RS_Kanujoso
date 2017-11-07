VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLaporanBulananPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Laporan Bulanan Jumlah PNS, CPNS & TPHL"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLaporanBulananPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5235
   Begin VB.CommandButton cmdtutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   2760
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
      Width           =   4755
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
         Width           =   3015
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MMMM yyyy "
            Format          =   117833731
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
            Format          =   117833731
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
      Picture         =   "frmLaporanBulananPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLaporanBulananPegawai.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanBulananPegawai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "frmLaporanBulananPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    On Error GoTo errLoad
    Dim pesan As VbMsgBoxResult
    strSQL = "select NamaRuangan, SUM(TPHLSD) as SD, SUM(TPHLSLTP) as SMP, SUM(TPHLSLTA) as SMA," & _
    "SUM(TPHLD3) as D3, SUM(TPHLS1) as S1, SUM(TPHLS2) as S2, SUM(TPHLLain2) as Lain2, " & _
    "SUM(CPNSIa) as CPNS1a, SUM(CPNSIc) as CPNS1c, SUM(CPNSIIa) as CPNS2a, SUM(CPNSIIb) as CPNS2b, " & _
    "SUM(CPNSIIc) as CPNS2c, SUM(CPNSIIIa) as CPNS3a, SUM(CPNSIIIb) as CPNS3b, " & _
    "SUM(PNSIa) as PNS1a, SUM(PNSIb) as PNS1b, SUM(PNSIc) as PNS1c, SUM(PNSId) as PNS1d, " & _
    "SUM(PNSIIa) as PNS2a, SUM(PNSIIb) as PNS2b, SUM(PNSIIc) as PNS2c, SUM(PNSIId) as PNS2d, " & _
    "SUM(PNSIIIa) as PNS3a, SUM(PNSIIIb) as PNS3b, SUM(PNSIIIc) as PNS3c, SUM(PNSIIId) as PNS3d, " & _
    "SUM(PNSIVa) as PNS4a, SUM(PNSIVb) as PNS4b, SUM(PNSIVc) as PNS4c, SUM(PNSIVd) as PNS4d, SUM(PNSIVe) as PNS4e " & _
    "From V_LaporanBulananJumlahPegawai where TglKeluar < '" & Format(DTPickerAwal, "yyyy/mm/28") & "'  OR TglKeluar is null " & _
    "group by NamaRuangan "
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        MsgBox "Tidak ada data  ", vbCritical, "Validasi"
        Exit Sub
    End If

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    frm_cetak_LaporanBulananPegawai.Show
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
