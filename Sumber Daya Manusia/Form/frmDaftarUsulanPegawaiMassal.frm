VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarUsulanPegawaiMassal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Riwayat Usulan & Realisasi Usulan Pegawai"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarUsulanPegawaiMassal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   15030
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   6960
      Width           =   15015
      Begin VB.CommandButton cmdDetail 
         Caption         =   "&Detail Usulan"
         Height          =   495
         Left            =   8880
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Realisasi Usulan"
         Height          =   495
         Left            =   8880
         TabIndex        =   19
         Top             =   18240
         Width           =   1935
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   210
         TabIndex        =   0
         Top             =   420
         Width           =   4335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12975
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdRealisasiBetul 
         Caption         =   "&Realisasi Usulan"
         Height          =   495
         Left            =   10920
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cari No. Riwayat  "
         Height          =   210
         Index           =   3
         Left            =   225
         TabIndex        =   11
         Top             =   180
         Width           =   1440
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
   Begin VB.Frame Frame2 
      Height          =   6015
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   15015
      Begin VB.Frame Frame4 
         Caption         =   "Filter Riwayat Usulan Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   8535
         Begin VB.OptionButton Option1 
            Caption         =   "Kenaikan Pangkat"
            Height          =   495
            Left            =   1560
            TabIndex        =   18
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Kenaikan Gaji"
            Height          =   495
            Left            =   1560
            TabIndex        =   17
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Pensiun"
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton option6 
            Caption         =   "Pengangkatan Pegawai Negeri Sipil"
            Height          =   495
            Left            =   3480
            TabIndex        =   15
            Top             =   240
            Width           =   3255
         End
         Begin VB.OptionButton option5 
            Caption         =   "Pengangkatan Pegawai TPHL"
            Height          =   495
            Left            =   3480
            TabIndex        =   14
            Top             =   600
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Periode Riwayat"
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
         Left            =   8760
         TabIndex        =   8
         Top             =   240
         Width           =   6135
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            Format          =   127533059
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3840
            TabIndex        =   4
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            Format          =   127533059
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
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
            Left            =   3480
            TabIndex        =   9
            Top             =   315
            Width           =   225
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
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
      Left            =   13200
      Picture         =   "frmDaftarUsulanPegawaiMassal.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarUsulanPegawaiMassal.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarUsulanPegawaiMassal.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "frmDaftarUsulanPegawaiMassal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long

Private Sub cmdDetail_Click()
    On Error GoTo hell
    If DataGrid1.ApproxCount = 0 Then Exit Sub
    If Len(Trim(DataGrid1.Columns("NoRiwayatRealisasi"))) > 0 Then
        MsgBox "Proses ini hanya untuk data riwayat usulan yang belum di realisasikan", vbInformation, "Informasi"
        Exit Sub
    End If
    Call subLoadFormRiwayatUsulan
    Exit Sub
hell:
End Sub

Private Sub subLoadFormRiwayatUsulan()
    On Error GoTo hell

    With frmUsulanRealisasi
        .Show
        .txtNamaFormPengirim = Me.Name
        .txtNoRiwayat.Text = DataGrid1.Columns("No.Riwayat").Value
        .subLoadDataUsulan
    End With

hell:
End Sub

Private Sub cmdRealisasiBetul_Click()
    If DataGrid1.ApproxCount = 0 Then Exit Sub
    If Len(Trim(DataGrid1.Columns("NoRiwayatRealisasi"))) > 0 Then
        MsgBox "Usulan sudah di Realisasikan ", vbInformation, "Informasi"
        Exit Sub
    End If
    Call subLoadFormRealisasi
End Sub

Private Sub subLoadFormRealisasi()
    On Error GoTo hell
    With frmUsulanRealisasiSetuju
        .Show
        .txtNoRiwayat.Text = DataGrid1.Columns("No.Riwayat").Value
        .subLoadDataUsulan
        .txtNoSK.SetFocus

    End With
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Public Sub cmdCari_Click()
    On Error GoTo errTampilkan

    Call subLoadDataGrid
    If DataGrid1.ApproxCount = 0 Then dtpAwal.SetFocus Else DataGrid1.SetFocus

    Exit Sub
errTampilkan:
End Sub

Private Sub DataGrid1_Click()
    WheelHook.WheelUnHook
    Set MyProperty = DataGrid1
    WheelHook.WheelHook DataGrid1
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdRealisasiBetul.SetFocus
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errFormLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    option1.Value = True
    Call cmdCari_Click

    Exit Sub
errFormLoad:
    msubPesanError
End Sub

Private Sub subLoadDataGrid()
    On Error GoTo errLoad
    Dim i As Integer
    '' kenaikan pangkat
    If option1.Value = True Then
        strSQL = "SELECT * FROM V_RiwayatSK " & _
        "WHERE [Tgl.Riwayat] BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        "AND [No.Riwayat] LIKE '%" & txtParameter.Text & "%' AND KdPangkatUsulan IS NOT NULL"

        Call msubRecFO(rs, strSQL)
        Set DataGrid1.DataSource = rs
        With DataGrid1
            For i = 0 To .Columns.Count - 1
                .Columns(i).Width = 0
            Next i
            .Columns("Tgl.Riwayat").Width = 1200
            .Columns("No.Riwayat").Width = 1200
            .Columns("Tgl.SK").Width = 1200
            .Columns("No.SK").Width = 2000
            .Columns("Penanda Tangan 1").Width = 2000
            .Columns("Penanda Tangan 2").Width = 2000
            .Columns("Keterangan").Width = 2500
            .Columns("NoRiwayatRealisasi").Width = 2000
        End With
    End If

    'kenaikan gaji
    If Option2.Value = True Then
        strSQL = "SELECT * FROM V_RiwayatSK " & _
        "WHERE [Tgl.Riwayat] BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        "AND [No.Riwayat] LIKE '%" & txtParameter.Text & "%' AND [Gaji Pokok Usulan] IS NOT NULL "

        msubRecFO rs, strSQL
        Set DataGrid1.DataSource = rs
        With DataGrid1
            For i = 0 To .Columns.Count - 1
                .Columns(i).Width = 0
            Next i
            .Columns("Tgl.Riwayat").Width = 1200
            .Columns("No.Riwayat").Width = 1200
            .Columns("Tgl.SK").Width = 1200
            .Columns("No.SK").Width = 2000
            .Columns("Penanda Tangan 1").Width = 2000
            .Columns("Penanda Tangan 2").Width = 2000
            .Columns("Keterangan").Width = 2500
            .Columns("NoRiwayatRealisasi").Width = 2000
        End With
    End If

    'pensiun
    If Option3.Value = True Then
        strSQL = "SELECT * FROM V_RiwayatSK " & _
        "WHERE [Tgl.Riwayat] BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        "AND [No.Riwayat] LIKE '%" & txtParameter.Text & "%' AND KdStatusUsulan ='05'"

        msubRecFO rs, strSQL
        Set DataGrid1.DataSource = rs
        With DataGrid1
            For i = 0 To .Columns.Count - 1
                .Columns(i).Width = 0
            Next i
            .Columns("Tgl.Riwayat").Width = 1200
            .Columns("No.Riwayat").Width = 1200
            .Columns("Tgl.SK").Width = 1200
            .Columns("No.SK").Width = 2000
            .Columns("Penanda Tangan 1").Width = 2000
            .Columns("Penanda Tangan 2").Width = 2000
            .Columns("Keterangan").Width = 2500
            .Columns("NoRiwayatRealisasi").Width = 2000
        End With
    End If

    'pengangkatan PNS
    If option6.Value = True Then
        strSQL = "SELECT * FROM V_RiwayatSK " & _
        "WHERE [Tgl.Riwayat] BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        "AND [No.Riwayat] LIKE '%" & txtParameter.Text & "%' AND KdDKategoryPUsulan IS NOT NULL"

        msubRecFO rs, strSQL
        Set DataGrid1.DataSource = rs
        With DataGrid1
            For i = 0 To .Columns.Count - 1
                .Columns(i).Width = 0
            Next i
            .Columns("Tgl.Riwayat").Width = 1200
            .Columns("No.Riwayat").Width = 1200
            .Columns("Tgl.SK").Width = 1200
            .Columns("No.SK").Width = 2000
            .Columns("Penanda Tangan 1").Width = 2000
            .Columns("Penanda Tangan 2").Width = 2000
            .Columns("Keterangan").Width = 2500
            .Columns("NoRiwayatRealisasi").Width = 2000
        End With
    End If

    'pengangkatan TPHL
    If option5.Value = True Then
        strSQL = "SELECT * FROM V_RiwayatSK " & _
        "WHERE [Tgl.Riwayat] BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        "AND [No.Riwayat] LIKE '%" & txtParameter.Text & "%' AND KdStatusUsulan IS NOT NULL"

        msubRecFO rs, strSQL
        Set DataGrid1.DataSource = rs
        With DataGrid1
            For i = 0 To .Columns.Count - 1
                .Columns(i).Width = 0
            Next i
            .Columns("Tgl.Riwayat").Width = 1200
            .Columns("No.Riwayat").Width = 1200
            .Columns("Tgl.SK").Width = 1200
            .Columns("No.SK").Width = 2000
            .Columns("Penanda Tangan 1").Width = 2000
            .Columns("Penanda Tangan 2").Width = 2000
            .Columns("Keterangan").Width = 2500
            .Columns("NoRiwayatRealisasi").Width = 2000
        End With
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtParameter_Change()
    Call subLoadDataGrid
    txtParameter.SetFocus: txtParameter.SelLength = Len(txtParameter.Text)
End Sub
