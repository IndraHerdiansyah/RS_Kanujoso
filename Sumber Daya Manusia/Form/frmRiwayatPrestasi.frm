VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatPrestasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Prestasi Pegawai"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmRiwayatPrestasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10695
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
      Left            =   5400
      TabIndex        =   6
      Top             =   6480
      Width           =   1215
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
      Left            =   9360
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
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
      Left            =   6720
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
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
      Left            =   8040
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Prestasi"
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
      TabIndex        =   10
      Top             =   1080
      Width           =   10455
      Begin VB.TextBox Text1 
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
         Left            =   6720
         MaxLength       =   100
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtPimpinanInstansiPemberi 
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
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtNamaPenghargaan 
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
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   1
         Top             =   600
         Width           =   5535
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
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         MaxLength       =   200
         TabIndex        =   5
         Top             =   2040
         Width           =   9975
      End
      Begin VB.TextBox txtNamaInstansiPemberi 
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
         TabIndex        =   3
         Top             =   1320
         Width           =   6735
      End
      Begin MSComCtl2.DTPicker dtpTglDiperoleh 
         Height          =   330
         Left            =   8040
         TabIndex        =   2
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
         Format          =   478150656
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Penghargaan"
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
         Left            =   6720
         TabIndex        =   20
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pimpinan Instansi Pemberi"
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
         Left            =   7080
         TabIndex        =   18
         Top             =   1080
         Width           =   1860
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Diperoleh"
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
         Left            =   8040
         TabIndex        =   15
         Top             =   360
         Width           =   930
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Penghargaan"
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
         Left            =   1080
         TabIndex        =   13
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label28 
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
         TabIndex        =   12
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Nama Instansi Pemberi"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1635
      End
   End
   Begin MSDataGridLib.DataGrid dgPrestasi 
      Height          =   2535
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4471
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
      TabIndex        =   17
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
      Picture         =   "frmRiwayatPrestasi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatPrestasi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPrestasi.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatPrestasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadPrestasi
    txtNamaPenghargaan.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoUrut.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPrestasi WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If Periksa("text", txtNamaPenghargaan, "Silahkan isi nama penghargaan ") = False Then Exit Sub
    If Periksa("text", txtNamaInstansiPemberi, "Silahkan isi nama instansi pemberi ") = False Then Exit Sub

    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Null)
        End If
        .Parameters.Append .CreateParameter("NamaPenghargaan", adVarChar, adParamInput, 100, Trim(txtNamaPenghargaan.Text))
        .Parameters.Append .CreateParameter("TglDiperoleh", adDate, adParamInput, , Format(dtpTglDiperoleh.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("InstansiPemberi", adVarChar, adParamInput, 100, Trim(txtNamaInstansiPemberi.Text))
        .Parameters.Append .CreateParameter("PimpinanInstansiPemberi", adVarChar, adParamInput, 30, IIf(txtPimpinanInstansiPemberi.Text = "", Null, Trim(txtPimpinanInstansiPemberi.Text)))
        .Parameters.Append .CreateParameter("NomorPiagam", adVarChar, adParamInput, 50, IIf(Text1.Text = "", Null, Trim(Text1.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adVarChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 2, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RPrestasi"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Riwayat Prestasi pegawai", vbCritical, "Validasi"
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
    Call subLoadPrestasi
    Call subClearData
    txtNamaPenghargaan.SetFocus
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Call frmRiwayatPegawai.subLoadRiwayatPrestasi
    Unload Me
End Sub

Private Sub dgPrestasi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgPrestasi.ApproxCount = 0 Then Exit Sub
    txtNoUrut.Text = dgPrestasi.Columns(1).Value
    txtNamaPenghargaan.Text = dgPrestasi.Columns(2).Value
    dtpTglDiperoleh.Value = dgPrestasi.Columns(3).Value
    If IsNull(dgPrestasi.Columns(4).Value) Then txtNamaInstansiPemberi.Text = "" Else txtNamaInstansiPemberi.Text = dgPrestasi.Columns(4).Value
    If IsNull(dgPrestasi.Columns(5).Value) Then txtPimpinanInstansiPemberi.Text = "" Else txtPimpinanInstansiPemberi.Text = dgPrestasi.Columns(5).Value
    If IsNull(dgPrestasi.Columns(6).Value) Then Text1.Text = "" Else Text1.Text = dgPrestasi.Columns(6).Value
    If IsNull(dgPrestasi.Columns(7).Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgPrestasi.Columns(7).Value
End Sub

Private Sub dtpTglDiperoleh_Change()
    dtpTglDiperoleh.MaxDate = Now
End Sub

Private Sub dtpTglDiperoleh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNamaInstansiPemberi.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadPrestasi
End Sub

Private Sub subLoadPrestasi()
    On Error GoTo hell
    strSQL = "SELECT * FROM RiwayatPrestasi WHERE IdPegawai='" & mstrIdPegawai & "' ORDER BY NoUrut"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgPrestasi
        Set .DataSource = rs
        .Columns("IdPegawai").Width = 0           'IdPegawai
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Caption = "No. Urut"
        .Columns("NamaPenghargaan").Width = 2500
        .Columns("NamaPenghargaan").Caption = "Nama Penghargaan"
        .Columns("TglDiperoleh").Width = 1500
        .Columns("TglDiperoleh").Caption = "Tgl. Diperoleh"
        .Columns("InstansiPemberi").Width = 2500
        .Columns("InstansiPemberi").Caption = "Instansi Pemberi"
        .Columns("PimpinanInstansiPemberi").Width = 1800
        .Columns("PimpinanInstansiPemberi").Caption = "Pimpinan Instansi"
        .Columns("Keterangan").Width = 3000
        .Columns("IdUser").Caption = "Nama User"
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    txtNamaPenghargaan.Text = ""
    dtpTglDiperoleh.Value = Format(Now, "dd/mm/yyyy")
    txtNamaInstansiPemberi.Text = ""
    txtPimpinanInstansiPemberi.Text = ""
    Text1.Text = ""
    txtKeterangan.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatPrestasi
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglDiperoleh.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNamaPenghargaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then Text1.SetFocus
End Sub

Private Sub dtpTahunDiperoleh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNamaInstansiPemberi.SetFocus
End Sub

Private Sub txtNamaInstansiPemberi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtPimpinanInstansiPemberi.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaPenghargaan_LostFocus()
    txtNamaPenghargaan.Text = StrConv(txtNamaPenghargaan, vbUpperCase)
End Sub

Private Sub txtPimpinanInstansiPemberi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub
