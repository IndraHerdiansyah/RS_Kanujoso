VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatOrganisasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Organisasi Pegawai"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmRiwayatOrganisasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10695
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
      Left            =   7800
      TabIndex        =   10
      Top             =   6720
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
      Left            =   6360
      TabIndex        =   9
      Top             =   6720
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
      Left            =   9240
      TabIndex        =   11
      Top             =   6720
      Width           =   1335
   End
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
      Left            =   4920
      TabIndex        =   8
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Organisasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   10455
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4080
         MaxLength       =   200
         TabIndex        =   7
         Top             =   2040
         Width           =   6135
      End
      Begin VB.TextBox txtAlamatOrganisasi 
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
         Left            =   4680
         MaxLength       =   200
         TabIndex        =   5
         Top             =   1320
         Width           =   5535
      End
      Begin VB.TextBox txtNamaPemimpinOrganisasi 
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
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2040
         Width           =   3735
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtJabatanOrganisasi 
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
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtNamaOrganisasi 
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
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   1
         Top             =   600
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker dtpTglMasuk 
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   129761280
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   330
         Left            =   2280
         TabIndex        =   4
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
         Format          =   129761280
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Akhir"
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
         Left            =   2280
         TabIndex        =   22
         Top             =   1080
         Width           =   615
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
         Left            =   4080
         TabIndex        =   21
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Organisasi"
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
         Left            =   4680
         TabIndex        =   18
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pimpinan Organisasi"
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
         Top             =   1800
         Width           =   1875
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Organisasi"
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
         Left            =   1440
         TabIndex        =   16
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan Organisasi"
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
         Left            =   6120
         TabIndex        =   15
         Top             =   360
         Width           =   1380
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
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Masuk"
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
         Top             =   1080
         Width           =   705
      End
   End
   Begin MSDataGridLib.DataGrid dgOrganisasi 
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4895
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
      TabIndex        =   20
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
      Picture         =   "frmRiwayatOrganisasi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatOrganisasi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatOrganisasi.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatOrganisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadOrganisasi
    txtNamaOrganisasi.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoUrut.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatOrganisasi WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If Periksa("text", txtNamaOrganisasi, "Silahkan isi nama organisasi ") = False Then Exit Sub
    If Periksa("text", txtJabatanOrganisasi, "Silahkan isi jabatan di organisasi ") = False Then Exit Sub

    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, Null)
        End If
        .Parameters.Append .CreateParameter("NamaOrganisasi", adVarChar, adParamInput, 100, Trim(txtNamaOrganisasi.Text))
        .Parameters.Append .CreateParameter("Jabatan", adVarChar, adParamInput, 50, Trim(txtJabatanOrganisasi.Text))
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglMasuk.Value, "yyyy/MM/dd"))
        If IsNull(dtpTglAkhir.Value) Then
            .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Format(dtpTglAkhir.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("AlamatOrganisasi", adVarChar, adParamInput, 200, IIf(txtAlamatOrganisasi.Text = "", Null, Trim(txtAlamatOrganisasi.Text)))
        .Parameters.Append .CreateParameter("PimpinanOrganisasi", adVarChar, adParamInput, 30, IIf(txtNamaPemimpinOrganisasi.Text = "", Null, Trim(txtNamaPemimpinOrganisasi.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 3, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_ROrganisasi"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Riwayat Organisasi pegawai", vbCritical, "Validasi"
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
    Call subLoadOrganisasi
    Call subClearData
    txtNamaOrganisasi.SetFocus
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgOrganisasi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgOrganisasi.ApproxCount = 0 Then Exit Sub
    txtNoUrut.Text = dgOrganisasi.Columns(1).Value
    txtNamaOrganisasi.Text = dgOrganisasi.Columns(2).Value
    txtJabatanOrganisasi.Text = dgOrganisasi.Columns(3).Value
    dtpTglMasuk.Value = dgOrganisasi.Columns(4).Value
    If IsNull(dgOrganisasi.Columns(5).Value) Then dtpTglAkhir.Value = Null Else dtpTglAkhir.Value = dgOrganisasi.Columns(5).Value
    If IsNull(dgOrganisasi.Columns(6).Value) Then txtAlamatOrganisasi.Text = "" Else txtAlamatOrganisasi.Text = dgOrganisasi.Columns(6).Value
    If IsNull(dgOrganisasi.Columns(7).Value) Then txtNamaPemimpinOrganisasi.Text = "" Else txtNamaPemimpinOrganisasi.Text = dgOrganisasi.Columns(7).Value
    If IsNull(dgOrganisasi.Columns(8).Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgOrganisasi.Columns(8).Value
End Sub

Private Sub dtpTglAkhir_Change()
    dtpTglAkhir.MaxDate = Now
End Sub

Private Sub dtpTglAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtAlamatOrganisasi.SetFocus
End Sub

Private Sub dtpTglMasuk_Change()
    On Error Resume Next
    dtpTglMasuk.MaxDate = Now
End Sub

Private Sub dtpTglMasuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadOrganisasi
End Sub

Private Sub subLoadOrganisasi()
    On Error GoTo hell
    strLSQL = "SELECT * FROM RiwayatOrganisasi WHERE IdPegawai='" & mstrIdPegawai & "' ORDER BY NoUrut"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgOrganisasi
        Set .DataSource = rs
        .Columns(0).Width = 0           'IdPegawai
        .Columns(1).Width = 1000
        .Columns(1).Caption = "No. Urut"
        .Columns(2).Width = 2000
        .Columns(2).Caption = "Organisasi"
        .Columns(3).Width = 1000
        .Columns(3).Caption = "Jabatan"
        .Columns(4).Width = 1000
        .Columns(4).Caption = "Tgl. Masuk"
        .Columns(5).Width = 1000
        .Columns(5).Caption = "Tgl. Akhir"
        .Columns(6).Width = 3000
        .Columns(6).Caption = "Alamat Organisasi"
        .Columns(7).Width = 3000
        .Columns(7).Caption = "Pimpinan Organisasi"
        .Columns("Keterangan").Width = 3500
        .Columns("IdUser").Caption = "Nama User"
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    txtNamaOrganisasi.Text = ""
    txtJabatanOrganisasi.Text = ""
    txtAlamatOrganisasi.Text = ""
    dtpTglMasuk.Value = Format(Now, "dd/mm/yyyy")
    dtpTglAkhir.Value = Format(Now, "dd/mm/yyyy")
    txtNamaPemimpinOrganisasi.Text = ""
    txtKeterangan.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatOrganisasi
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaOrganisasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtJabatanOrganisasi.SetFocus
End Sub

Private Sub txtJabatanOrganisasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglMasuk.SetFocus
End Sub

Private Sub txtAlamatOrganisasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNamaPemimpinOrganisasi.SetFocus
End Sub

Private Sub txtNamaPemimpinOrganisasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub
