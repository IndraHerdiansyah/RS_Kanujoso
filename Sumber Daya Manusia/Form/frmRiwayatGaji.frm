VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatGaji 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Gaji Pegawai"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   Icon            =   "frmRiwayatGaji.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10935
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
      Left            =   6600
      TabIndex        =   10
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
      Left            =   9480
      TabIndex        =   12
      Top             =   6840
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
      Left            =   5160
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Gaji"
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
      TabIndex        =   14
      Top             =   1080
      Width           =   10695
      Begin VB.TextBox txtMasaKerja 
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
         Left            =   120
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
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
         Left            =   120
         MaxLength       =   200
         TabIndex        =   8
         Top             =   2040
         Width           =   10335
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
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1320
         Width           =   2775
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
         Left            =   960
         MaxLength       =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox txtNoUrut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtJumlah 
         Alignment       =   1  'Right Justify
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
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "0"
         Top             =   1320
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpTglSK 
         Height          =   330
         Left            =   5760
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
         Format          =   121241600
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSDataListLib.DataCombo dcKdKomponenGaji 
         Height          =   315
         Left            =   5760
         TabIndex        =   6
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
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
      Begin MSComCtl2.DTPicker dtpTglBerlaku 
         Height          =   330
         Left            =   8160
         TabIndex        =   3
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
         Format          =   120979456
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. SK"
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
         Left            =   5760
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Berlaku "
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
         Left            =   8160
         TabIndex        =   23
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Masa Kerja Golongan"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label10 
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
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label8 
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
         Left            =   2880
         TabIndex        =   20
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label4 
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
         Left            =   960
         TabIndex        =   19
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Komponen Gaji"
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
         Left            =   5760
         TabIndex        =   16
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
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
         Left            =   8160
         TabIndex        =   15
         Top             =   1080
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid dgGaji 
      Height          =   2895
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5106
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
      Left            =   9120
      Picture         =   "frmRiwayatGaji.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatGaji.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatGaji.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Function sp_RiwayatGaji() As Boolean
    On Error GoTo errSimpan
    sp_RiwayatGaji = True
    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Null)
        End If
        .Parameters.Append .CreateParameter("NoSK", adVarChar, adParamInput, 30, IIf(txtNoSK.Text = "", Null, Trim(txtNoSK.Text)))
        If IsNull(dtpTglSK.Value) Then
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Format(dtpTglSK.Value, "yyyy/MM/dd"))
        End If
        If IsNull(dtpTglBerlaku.Value) Then
            .Parameters.Append .CreateParameter("TglBerlakuSK", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglBerlakuSK", adDate, adParamInput, , Format(dtpTglBerlaku.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("MasaKerja", adVarChar, adParamInput, 50, IIf(txtMasaKerja.Text = "", Null, Trim(txtMasaKerja.Text)))
        .Parameters.Append .CreateParameter("TandaTanganSK", adVarChar, adParamInput, 30, IIf(txtTandaTanganSK.Text = "", Null, Trim(txtTandaTanganSK.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 2, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RGaji"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Riwayat Gaji Pegawai ", vbCritical, "Validasi"
            Set adoCommand = Nothing
            sp_RiwayatGaji = False
        Else
            txtNoUrut.Text = .Parameters("OutputNoUrut").Value

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errSimpan:
    msubPesanError
End Function

Private Function sp_RiwayatDetailGaji() As Boolean
    On Error GoTo errSimpan
    sp_RiwayatDetailGaji = True
    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, txtNoUrut.Text)
        .Parameters.Append .CreateParameter("KdKomponenGaji", adChar, adParamInput, 2, dcKdKomponenGaji.BoundText)
        .Parameters.Append .CreateParameter("Jumlah", adCurrency, adParamInput, , CCur(txtJumlah.Text))

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_DetailRGaji"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Detail Riwayat Gaji pegawai", vbCritical, "Validasi"
            Set adoCommand = Nothing
            sp_RiwayatDetailGaji = False
        Else
            MsgBox "Data berhasi disimpan ", vbInformation, "Informasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errSimpan:
    msubPesanError
End Function

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadGaji
    txtNoSK.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If dgGaji.ApproxCount = 0 Then Exit Sub

    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM DetailRiwayatGaji WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut ='" & txtNoUrut.Text & "' AND KdKomponenGaji = '" & dcKdKomponenGaji.BoundText & "'  "
    dbConn.Execute strSQL
    strSQL = "DELETE FROM RiwayatGaji WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut ='" & txtNoUrut.Text & "' "
    dbConn.Execute strSQL
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If dcKdKomponenGaji.Text <> "" Then
        If Periksa("datacombo", dcKdKomponenGaji, "Komponen Gaji Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If Periksa("datacombo", dcKdKomponenGaji, "Silahkan isi komponen gaji ") = False Then Exit Sub
    If Periksa("text", txtJumlah, "Silahkan isi jumlah gaji ") = False Then Exit Sub

    If sp_RiwayatGaji() = False Then Exit Sub
    If sp_RiwayatDetailGaji() = False Then Exit Sub

    Call cmdBatal_Click
    dcKdKomponenGaji.SetFocus
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Call frmRiwayatPegawai.subLoadRiwayatGaji
    Unload Me
End Sub

Private Sub dcKdKomponenGaji_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKdKomponenGaji.Text)) = 0 Then txtJumlah.SetFocus: Exit Sub
        If dcKdKomponenGaji.MatchedWithList = True Then txtJumlah.SetFocus: Exit Sub
        strSQL = "Select KdKomponenGaji,KomponenGaji from KomponenGaji WHERE (KomponenGaji LIKE '%" & dcKdKomponenGaji.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKdKomponenGaji.BoundText = rs(0).Value
        dcKdKomponenGaji.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgGaji_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgGaji.ApproxCount = 0 Then Exit Sub
    txtNoUrut.Text = dgGaji.Columns("NoUrut").Value
    If IsNull(dgGaji.Columns("NoSK").Value) Then txtNoSK.Text = "" Else txtNoSK.Text = dgGaji.Columns("NoSK").Value
    If IsNull(dgGaji.Columns("TglSK").Value) Then dtpTglSK.Value = Null Else dtpTglSK.Value = dgGaji.Columns("TglSK").Value
    If IsNull(dgGaji.Columns("TglBerlakuSK").Value) Then dtpTglBerlaku.Value = Null Else dtpTglBerlaku.Value = dgGaji.Columns("TglBerlakuSK").Value
    If IsNull(dgGaji.Columns("MasaKerja").Value) Then txtMasaKerja.Text = "" Else txtMasaKerja.Text = dgGaji.Columns("MasaKerja").Value
    If IsNull(dgGaji.Columns("TandaTanganSK").Value) Then txtTandaTanganSK.Text = "" Else txtTandaTanganSK.Text = dgGaji.Columns("TandaTanganSK").Value
    dcKdKomponenGaji.BoundText = dgGaji.Columns("KdKomponenGaji").Value
    txtJumlah.Text = dgGaji.Columns("Jumlah").Value
    Call txtJumlah_LostFocus
    If IsNull(dgGaji.Columns("Keterangan").Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgGaji.Columns("Keterangan").Value
End Sub

Private Sub dtpTglBerlaku_Change()
    dtpTglBerlaku.MaxDate = Now
End Sub

Private Sub dtpTglBerlaku_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtMasaKerja.SetFocus
End Sub

Private Sub dtpTglSK_Change()
    dtpTglSK.MaxDate = Now
End Sub

Private Sub dtpTglSK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglBerlaku.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call SetComboKomponenGaji
    Call subLoadGaji
End Sub

Private Sub subLoadGaji()
    On Error GoTo hell
    strLSQL = "SELECT IdPegawai,NoUrut, NoSK, TglSK, TglBerlakuSK, MasaKerja, TandaTanganSK, KomponenGaji, Jumlah, Keterangan, NamaUser, KdKomponenGaji " & _
    " FROM V_RiwayatGaji WHERE IdPegawai='" & mstrIdPegawai & "' order by NoUrut"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgGaji
        Set .DataSource = rs
        .Columns("IdPegawai").Width = 0           'IdPegawai
        .Columns("NoUrut").Width = 1000
        .Columns("TglSK").Width = 1200
        .Columns("TglBerlakuSK").Width = 1200
        .Columns("MasaKerja").Width = 1500
        .Columns("TandaTanganSK").Width = 2000
        .Columns("KomponenGaji").Width = 2000
        .Columns("Jumlah").Width = 2200
        .Columns("Jumlah").NumberFormat = "#,##"
        .Columns("Keterangan").Width = 3800
        .Columns("NamaUser").Width = 2000
        .Columns("KdKomponenGaji").Width = 0
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    txtNoSK.Text = ""
    txtTandaTanganSK.Text = ""
    dcKdKomponenGaji.Text = ""
    txtJumlah.Text = 0
    txtKeterangan.Text = ""
    txtMasaKerja.Text = ""
    dtpTglSK.Value = Format(Now, "dd/mmmm/yyyy")
    dtpTglBerlaku.Value = Format(Now, "dd/mmmm/yyyy")
    txtNoSK.SetFocus
End Sub

Private Sub dcKdKomponenGaji_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtJumlah.SetFocus
End Sub

Sub SetComboKomponenGaji()
    Set rs = Nothing
    strSQL = "Select * from KomponenGaji where StatusEnabled='1'"
    Call msubDcSource(dcKdKomponenGaji, rs, strSQL)
    Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatGaji
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtJumlah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtJumlah_LostFocus()
    txtJumlah.Text = IIf(Val(txtJumlah) = 0, 0, Format(txtJumlah, "#,###"))
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtMasaKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTandaTanganSK.SetFocus
End Sub

Private Sub txtNoSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglSK.SetFocus
End Sub

Private Sub txtTandaTanganSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKdKomponenGaji.SetFocus
End Sub
