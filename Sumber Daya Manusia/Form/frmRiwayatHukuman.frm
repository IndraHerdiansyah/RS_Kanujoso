VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatHukuman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Hukuman Pegawai"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmRiwayatHukuman.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Hukuman"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   9015
      Begin MSDataListLib.DataCombo dcJenisHukuman 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VB.TextBox txtKeterangan 
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
         Height          =   330
         Left            =   240
         MaxLength       =   200
         TabIndex        =   7
         Top             =   2040
         Width           =   8535
      End
      Begin VB.TextBox txtNoUrut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoSK 
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
         Height          =   330
         Left            =   4320
         MaxLength       =   30
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtTandaTanganSK 
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
         Height          =   330
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1320
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker dtpTglSK 
         Height          =   330
         Left            =   6600
         TabIndex        =   3
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   121176064
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin MSComCtl2.DTPicker DTPTglSelesai 
         Height          =   330
         Left            =   6600
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   121176064
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin MSComCtl2.DTPicker dtpTMT 
         Height          =   330
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   121176064
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Selesai"
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
         Left            =   6600
         TabIndex        =   22
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Mulai"
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
         TabIndex        =   21
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Left            =   6600
         TabIndex        =   20
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label15 
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
         TabIndex        =   19
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. Urut"
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
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Left            =   4320
         TabIndex        =   16
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Hukuman"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label12 
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
         Left            =   2520
         TabIndex        =   14
         Top             =   1080
         Width           =   1260
      End
   End
   Begin MSDataGridLib.DataGrid dgRiwayatHukuman 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      TabIndex        =   18
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
      Left            =   7440
      Picture         =   "frmRiwayatHukuman.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatHukuman.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatHukuman.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatHukuman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadRiwayatHukuman
    dcJenisHukuman.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatHukuman WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If dcJenisHukuman.Text <> "" Then
        If Periksa("datacombo", dcJenisHukuman, "Jenis Hukuman Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If Periksa("datacombo", dcJenisHukuman, "Silahkan isi jenis hukuman ") = False Then Exit Sub

    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Null)
        End If
        .Parameters.Append .CreateParameter("KdJenisHukuman", adVarChar, adParamInput, 3, dcJenisHukuman.BoundText)
        .Parameters.Append .CreateParameter("NoSK", adVarChar, adParamInput, 30, IIf(txtNoSK.Text = "", Null, txtNoSK.Text))
        If IsNull(dtpTglSK.Value) Then
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Format(dtpTglSK.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("TandaTanganSK", adVarChar, adParamInput, 30, IIf(txtTandaTanganSK.Text = "", Null, Trim(txtTandaTanganSK.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        If IsNull(DTPTglSelesai.Value) Then
            .Parameters.Append .CreateParameter("TglSelesai", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglSelesai", adDate, adParamInput, , Format(DTPTglSelesai.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        If IsNull(dtpTMT.Value) Then
            .Parameters.Append .CreateParameter("TMT", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TMT", adDate, adParamInput, , Format(dtpTMT.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 2, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RHukuman"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Riwayat Hukuman", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            Exit Sub
        Else
            txtNoUrut.Text = .Parameters("OutputNoUrut").Value
            MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Call subLoadRiwayatHukuman
    Call subClearData
    dcJenisHukuman.SetFocus
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisHukuman_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtNoSK.SetFocus

On Error GoTo Errload
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcJenisHukuman.Text)) = 0 Then txtNoSK.SetFocus: Exit Sub
        If dcJenisHukuman.MatchedWithList = True Then txtNoSK.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdJenisHukuman, JenisHukuman FROM JenisHukuman WHERE JenisHukuman LIKE '%" & dcJenisHukuman.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        dcJenisHukuman.BoundText = dbRst(0).Value
        dcJenisHukuman.Text = dbRst(1).Value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dgRiwayatHukuman_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgRiwayatHukuman.ApproxCount = 0 Then Exit Sub
    txtNoUrut.Text = dgRiwayatHukuman.Columns(1).Value
    dcJenisHukuman.Text = dgRiwayatHukuman.Columns(2).Value
    If IsNull(dgRiwayatHukuman.Columns(3).Value) Then txtNoSK.Text = "" Else txtNoSK.Text = dgRiwayatHukuman.Columns(3).Value
    If IsNull(dgRiwayatHukuman.Columns(4).Value) Then dtpTglSK.Value = Null Else dtpTglSK.Value = dgRiwayatHukuman.Columns(4).Value
    If IsNull(dgRiwayatHukuman.Columns(5).Value) Then txtTandaTanganSK.Text = "" Else txtTandaTanganSK.Text = dgRiwayatHukuman.Columns(5).Value
    If IsNull(dgRiwayatHukuman.Columns(6).Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgRiwayatHukuman.Columns(6).Value
    If IsNull(dgRiwayatHukuman.Columns(7).Value) Then DTPTglSelesai.Value = Null Else DTPTglSelesai.Value = dgRiwayatHukuman.Columns(7).Value
    If IsNull(dgRiwayatHukuman.Columns(9).Value) Then dtpTMT.Value = Null Else dtpTMT.Value = dgRiwayatHukuman.Columns(9).Value
End Sub

Private Sub DTPTglSelesai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub dtpTglSK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTMT.SetFocus
End Sub

Private Sub dtpTMT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTandaTanganSK.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTandaTanganSK.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadRiwayatHukuman
    Call subLoadDcSource
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatHukuman
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub subLoadRiwayatHukuman()
    On Error GoTo hell
    strLSQL = " SELECT * " & _
    " FROM v_RiwayatHukuman WHERE IdPegawai ='" & mstrIdPegawai & "' ORDER BY NoUrut"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgRiwayatHukuman
        Set .DataSource = rs
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Caption = "No. Urut"
        .Columns("JenisHukuman").Width = 1300
        .Columns("JenisHukuman").Caption = "Jenis Hukuman"
        .Columns("NoSK").Width = 2500
        .Columns("NoSK").Caption = "No. SK"
        .Columns("TglSK").Width = 1500
        .Columns("TglSK").Caption = "Tgl. SK"
        .Columns("TandaTanganSK").Width = 1700
        .Columns("TandaTanganSK").Caption = "TTD SK"
        .Columns("Keterangan").Width = 3200
        .Columns("TglSelesai").Width = 1500
        .Columns("TglSelesai").Caption = "Tgl. Selesai"
        .Columns("NamaUser").Width = 2200
        .Columns("NamaUser").Caption = "Nama User"
        .Columns("TMT").Width = 1500
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    strSQL = " SELECT KdJenisHukuman, JenisHukuman FROM JenisHukuman where statusenabled='1'"
    Call msubDcSource(dcJenisHukuman, rs, strSQL)
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    dcJenisHukuman.Text = ""
    txtNoSK.Text = ""
    dtpTglSK.Value = Format(Now, "dd/mm/yyyy")
    txtTandaTanganSK.Text = ""
    DTPTglSelesai.Value = Format(Now, "dd/mm/yyyy")
    txtKeterangan.Text = ""
    dtpTMT.Value = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub txtNoSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglSK.SetFocus
End Sub

Private Sub txtTandaTanganSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DTPTglSelesai.SetFocus
End Sub

