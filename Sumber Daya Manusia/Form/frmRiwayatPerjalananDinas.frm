VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatPerjalananDinas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Perjalanan Dinas Pegawai"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   Icon            =   "frmRiwayatPerjalananDinas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11055
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
      TabIndex        =   11
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdcetaksuratlangsung 
      Caption         =   "&Cetak Surat"
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
      Left            =   3600
      TabIndex        =   9
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   5280
      TabIndex        =   10
      Top             =   7440
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
      Left            =   9600
      TabIndex        =   13
      Top             =   7440
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
      Left            =   8160
      TabIndex        =   12
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Perjalanan Dinas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   10815
      Begin VB.TextBox txtkendaraan 
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
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2040
         Width           =   2415
      End
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
         MaxLength       =   200
         TabIndex        =   8
         Top             =   2760
         Width           =   10455
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
         MaxLength       =   3
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtKotaTuj 
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
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtTujuanKunj 
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
         Height          =   1050
         Left            =   240
         MaxLength       =   200
         TabIndex        =   3
         Top             =   1320
         Width           =   4695
      End
      Begin VB.TextBox txtNegTuj 
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
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtPenyandangDana 
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
         Left            =   7680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2040
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtpTglPergi 
         Height          =   330
         Left            =   5160
         TabIndex        =   4
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
         Format          =   129236992
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin MSComCtl2.DTPicker dtpTglPulang 
         Height          =   330
         Left            =   7200
         TabIndex        =   5
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
         Format          =   129236992
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Pulang"
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
         Left            =   7200
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kendaraan "
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
         Left            =   5160
         TabIndex        =   24
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label Label5 
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
         TabIndex        =   23
         Top             =   2520
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
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kota Tujuan"
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
         Left            =   1200
         TabIndex        =   20
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Negara Tujuan"
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
         Left            =   4920
         TabIndex        =   19
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tujuan Kunjungan"
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
         Width           =   1305
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Pergi"
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
         Left            =   5160
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Penyandang Dana"
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
         Left            =   7680
         TabIndex        =   16
         Top             =   1800
         Width           =   1320
      End
   End
   Begin MSDataGridLib.DataGrid dgPerjalanan 
      Height          =   2895
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   10815
      _ExtentX        =   19076
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
      TabIndex        =   22
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
      Left            =   9240
      Picture         =   "frmRiwayatPerjalananDinas.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatPerjalananDinas.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPerjalananDinas.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatPerjalananDinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadPerjalanan
    txtKotaTuj.SetFocus
End Sub

Private Sub cmdcetaksuratlangsung_Click()
    If txtNoUrut.Text = "" Then Exit Sub
    strSQL = "select * from V_CetakSuratPerjalananDinasPegawai where idpegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    Call msubRecFO(rs, strSQL)
    frmCetakSuratKeteranganPerjalananDinas.Show
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoUrut.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPerjalananDinas WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan

    If Periksa("text", txtKotaTuj, "Silahkan isi kota tujuan ") = False Then Exit Sub
    If Periksa("text", txtNegTuj, "Silahkan isi negara tujuan ") = False Then Exit Sub
    If Periksa("text", txtTujuanKunj, "Tujuan Kunjungan harus diisi!") = False Then Exit Sub

    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, Null)
        End If
        .Parameters.Append .CreateParameter("KotaTujuan", adVarChar, adParamInput, 50, Trim(txtKotaTuj.Text))
        .Parameters.Append .CreateParameter("NegaraTujuan", adVarChar, adParamInput, 50, Trim(txtNegTuj.Text))
        .Parameters.Append .CreateParameter("TujuanKunjungan", adVarChar, adParamInput, 200, Trim(txtTujuanKunj.Text))
        .Parameters.Append .CreateParameter("TglPergi", adDate, adParamInput, , Format(dtpTglPergi.Value, "yyyy/MM/dd"))
        If IsNull(dtpTglPulang.Value) Then
            .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , Format(dtpTglPulang.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("Kendaraan", adVarChar, adParamInput, 50, IIf(txtkendaraan.Text = "", Null, Trim(txtkendaraan.Text)))
        .Parameters.Append .CreateParameter("PenyandangDana", adVarChar, adParamInput, 50, IIf(txtPenyandangDana.Text = "", Null, Trim(txtPenyandangDana.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 3, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RPjlnDns"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Riwayat Perjalanan Dinas", vbCritical, "Validasi"
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
    Call subLoadPerjalanan
    Call subClearData
    Exit Sub
errSimpan:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
    frmRiwayatPegawai.Enabled = True
    Call frmRiwayatPegawai.subLoadRiwayatPerjalananDinas
End Sub

Private Sub dgPerjalanan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgPerjalanan.ApproxCount = 0 Then Exit Sub
    cmdHapus.Enabled = True
    cmdSimpan.Enabled = True
    txtNoUrut.Text = dgPerjalanan.Columns(1).Value
    txtKotaTuj.Text = dgPerjalanan.Columns(2).Value
    txtNegTuj.Text = dgPerjalanan.Columns(3).Value
    txtTujuanKunj.Text = dgPerjalanan.Columns(4).Value
    dtpTglPergi.Value = dgPerjalanan.Columns(5).Value
    If IsNull(dgPerjalanan.Columns(6).Value) Then dtpTglPulang.Value = Null Else dtpTglPulang.Value = dgPerjalanan.Columns(6).Value
    If IsNull(dgPerjalanan.Columns(7).Value) Then txtkendaraan.Text = "" Else txtkendaraan.Text = dgPerjalanan.Columns(7).Value
    If IsNull(dgPerjalanan.Columns(8).Value) Then txtPenyandangDana.Text = "" Else txtPenyandangDana.Text = dgPerjalanan.Columns(8).Value
    If IsNull(dgPerjalanan.Columns(9).Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgPerjalanan.Columns(9).Value
End Sub

Private Sub dtpTglPergi_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then dtpTglPulang.SetFocus
End Sub

Private Sub dtpTglPulang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtkendaraan.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadPerjalanan
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmRiwayatPegawai.Enabled = True
    Call frmRiwayatPegawai.subLoadRiwayatPerjalananDinas
End Sub

Private Sub txtkendaraan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPenyandangDana.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKotaTuj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNegTuj.SetFocus
End Sub

Private Sub txtLamaKunj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtPenyandangDana.SetFocus
End Sub

Private Sub txtNegTuj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtTujuanKunj.SetFocus
End Sub

Private Sub txtPenyandangDana_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtTujuanKunj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglPergi.SetFocus
End Sub

Private Sub subLoadPerjalanan()
    On Error GoTo hell
    strLSQL = "SELECT * FROM RiwayatPerjalananDinas WHERE IdPegawai ='" & mstrIdPegawai & "' ORDER BY NoUrut"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgPerjalanan
        Set .DataSource = rs
        .Columns("IdPegawai").Width = 0           'IdPegawai
        .Columns("NoUrut").Width = 800
        .Columns("NoUrut").Caption = "No. Urut"
        .Columns("KotaTujuan").Width = 1500
        .Columns("KotaTujuan").Caption = "Kota Tujuan"
        .Columns("NegaraTujuan").Width = 1400
        .Columns("NegaraTujuan").Caption = "Negara Tujuan"
        .Columns("TujuanKunjungan").Width = 1600
        .Columns("TujuanKunjungan").Caption = "Tujuan Kunjungan"
        .Columns("TglPergi").Width = 1200
        .Columns("TglPergi").Caption = "Tgl. Pergi"
        .Columns("TglPulang").Width = 1200
        .Columns("TglPulang").Caption = "Tgl. Pulang"
        .Columns("PenyandangDana").Width = 2000
        .Columns("PenyandangDana").Caption = "Penyandang Dana"
        .Columns("Keterangan").Width = 2000
        .Columns("IdUser").Caption = "Nama User"

    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    txtKotaTuj.Text = ""
    txtNegTuj.Text = "INDONESIA"
    txtTujuanKunj.Text = ""
    dtpTglPergi.Value = Format(Now, "dd/MM/yyyy")
    dtpTglPulang.Value = Format(Now, "dd/MM/yyyy")
    txtPenyandangDana.Text = ""
    txtKeterangan.Text = ""
    txtkendaraan.Text = ""
End Sub
