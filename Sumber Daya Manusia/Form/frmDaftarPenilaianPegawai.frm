VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaftarPenilaianPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Penilaian Pegawai"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPenilaianPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   14790
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   7440
      Width           =   14775
      Begin VB.CommandButton cmdUbah 
         Caption         =   "&Ubah"
         Height          =   495
         Left            =   11400
         TabIndex        =   15
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   495
         Left            =   9720
         TabIndex        =   14
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox txtparameter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         MaxLength       =   50
         TabIndex        =   12
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmdCetak1 
         Caption         =   "Cetak Lembar &1"
         Height          =   495
         Left            =   5880
         TabIndex        =   11
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton cmdCetak2 
         Caption         =   "Cetak Lembar &2"
         Height          =   495
         Left            =   7560
         TabIndex        =   10
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13080
         TabIndex        =   0
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cari Nama Pegawai"
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1635
      End
   End
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
      Height          =   6495
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   14775
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
         Height          =   855
         Left            =   10440
         TabIndex        =   3
         Top             =   120
         Width           =   4215
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
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
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   5520
            TabIndex        =   5
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   118358019
            UpDown          =   -1  'True
            CurrentDate     =   37967
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   1320
            TabIndex        =   6
            Top             =   300
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   118358019
            UpDown          =   -1  'True
            CurrentDate     =   37967
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   5160
            TabIndex        =   4
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPenilaianPegawai 
         Height          =   5295
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   3
         RowHeight       =   19
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
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
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
      Left            =   12960
      Picture         =   "frmDaftarPenilaianPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPenilaianPegawai.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPenilaianPegawai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmDaftarPenilaianPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdCari_Click()
    On Error GoTo hell
    Set rs = Nothing
    strSQL = "select * from V_DaftarPenilaianPegawai where namalengkap like '%" & txtParameter.Text & "%' and year(TglAkhir)= '" & Format(dtpAwal, "yyyy") & "'"
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPenilaianPegawai.DataSource = rs
    Set rs = Nothing
    lblJumData.Caption = "Data " & dgDaftarPenilaianPegawai.Bookmark & "/" & dgDaftarPenilaianPegawai.ApproxCount
    'Call subSetGrid
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdCetak1_Click()
    On Error GoTo hell
    If dgDaftarPenilaianPegawai.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    strSQL = "Select * from V_DaftarPenilaianPegawai where IdPegawai='" & dgDaftarPenilaianPegawai.Columns("IdPegawai").Value & "' and NoUrut='" & dgDaftarPenilaianPegawai.Columns("NoUrut").Value & "'"
    Call msubRecFO(rs, strSQL)
    frmCetakPenilaianPegawai.Show
    Exit Sub
hell:
End Sub

Private Sub cmdCetak2_Click()
    On Error GoTo hell
    If dgDaftarPenilaianPegawai.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    strSQL = "Select * from V_DaftarPenilaianPegawai where IdPegawai='" & dgDaftarPenilaianPegawai.Columns("IdPegawai").Value & "' and NoUrut='" & dgDaftarPenilaianPegawai.Columns("NoUrut").Value & "'"
    Call msubRecFO(rs, strSQL)
    frmCetakPenilaianPegawaiKeDua.Show
    Exit Sub
hell:
End Sub

Private Sub cmdTambah_Click()
    frmDataPerhitunganNilai.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdUbah_Click()
    If dgDaftarPenilaianPegawai.ApproxCount = 0 Then Exit Sub
    With frmDataPerhitunganNilai
        .txtNoUrut.Text = dgDaftarPenilaianPegawai.Columns("NoUrut").Value
        .txtKdPeg1.Text = dgDaftarPenilaianPegawai.Columns("IdPegawai").Value
        .txtNamaPegawai.Text = dgDaftarPenilaianPegawai.Columns("NamaLengkap").Value

        If dgDaftarPenilaianPegawai.Columns("NamaJabatan") = "" Then
            .txtJabatan.Text = ""
        Else
            .txtJabatan.Text = dgDaftarPenilaianPegawai.Columns("NamaJabatan").Value
        End If

        .txtKdPeg2.Text = dgDaftarPenilaianPegawai.Columns("IdPegawai2").Value
        .txtPenilai.Text = dgDaftarPenilaianPegawai.Columns("Penilai").Value
        .txtJabatanPenilai.Text = IIf(IsNull(dgDaftarPenilaianPegawai.Columns("JabatanPenilai")), "", (dgDaftarPenilaianPegawai.Columns("JabatanPenilai")))
        .txtKdPeg3.Text = dgDaftarPenilaianPegawai.Columns("IdPegawai3").Value
        .txtAtasanPenilai.Text = dgDaftarPenilaianPegawai.Columns("AtasanPenilai").Value
        .txtJabatanAtasan.Text = IIf(IsNull(dgDaftarPenilaianPegawai.Columns("JabatanAtasan")), "", (dgDaftarPenilaianPegawai.Columns("JabatanAtasan")))
        .txtKesetiaan.Text = dgDaftarPenilaianPegawai.Columns("NilaiKesetiaan").Value
        .txtPrestasi.Text = dgDaftarPenilaianPegawai.Columns("NilaiPrestasi").Value
        .txtTanggungJawab.Text = dgDaftarPenilaianPegawai.Columns("NilaiTanggungJawab").Value
        .txtKetaatan.Text = dgDaftarPenilaianPegawai.Columns("NilaiKetaatan").Value
        .txtKejujuran.Text = dgDaftarPenilaianPegawai.Columns("NilaiKejujuran").Value
        .txtKerjasama.Text = dgDaftarPenilaianPegawai.Columns("NilaiKerjasama").Value
        .txtPrakarsa.Text = dgDaftarPenilaianPegawai.Columns("NilaiPrakarsa").Value
        .txtKepemimpinan.Text = dgDaftarPenilaianPegawai.Columns("NilaiKepemimpinan").Value
        .fraPegawai3.Visible = False
    End With

End Sub

Private Sub dgDaftarPenilaianPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarPenilaianPegawai
    WheelHook.WheelHook dgDaftarPenilaianPegawai
End Sub

Private Sub dgDaftarPenilaianPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    lblJumData.Caption = "Data " & dgDaftarPenilaianPegawai.Bookmark & "/" & dgDaftarPenilaianPegawai.ApproxCount
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    If mblnAdmin = False Then
        cmdUbah.Enabled = False
    Else
        cmdUbah.Enabled = True
    End If
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00")
    dtpAkhir.Value = Now

    Call cmdCari_Click

End Sub

Private Sub txtParameter_Change()
    Call cmdCari_Click
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
