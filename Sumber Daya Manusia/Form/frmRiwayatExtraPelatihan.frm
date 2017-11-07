VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatExtraPelatihan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Extra Pelatihan Pegawai"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "frmRiwayatExtraPelatihan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10710
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
      TabIndex        =   10
      Top             =   7560
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
      TabIndex        =   13
      Top             =   7560
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
      TabIndex        =   11
      Top             =   7560
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
      Left            =   7800
      TabIndex        =   12
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Extra Pelatihan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   10455
      Begin VB.TextBox txtlamawaktu 
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
         Left            =   8640
         MaxLength       =   100
         TabIndex        =   28
         Top             =   1320
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dcjnspendidikan 
         Height          =   315
         Left            =   1080
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3720
         MaxLength       =   100
         TabIndex        =   9
         Top             =   2760
         Width           =   6495
      End
      Begin VB.TextBox txtPimpinanPenyelenggara 
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
         TabIndex        =   8
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtNamaPelatihan 
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
         Left            =   3720
         MaxLength       =   100
         TabIndex        =   2
         Top             =   600
         Width           =   6495
      End
      Begin VB.TextBox txtKedudukanPeranan 
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
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   3375
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
         Width           =   735
      End
      Begin VB.TextBox txtInstansiPenyelenggara 
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
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtAlamatPenyelenggaraan 
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
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2040
         Width           =   6495
      End
      Begin MSComCtl2.DTPicker dtpTglMulai 
         Height          =   330
         Left            =   3720
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
         CustomFormat    =   "yyyy"
         Format          =   45547520
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   330
         Left            =   6120
         TabIndex        =   5
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
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
         CustomFormat    =   "yyyy"
         Format          =   45547520
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Lama Waktu"
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
         Left            =   8640
         TabIndex        =   27
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label3 
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
         Left            =   6120
         TabIndex        =   26
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Jenis Pendidikan"
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
         Left            =   1080
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
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
         Left            =   3720
         TabIndex        =   24
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pimpinan Penyelenggara"
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
         Width           =   1755
      End
      Begin VB.Label Label1 
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
         Left            =   3720
         TabIndex        =   22
         Top             =   1080
         Width           =   615
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
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kedudukan/Peranan"
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
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelatihan"
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
         Left            =   3720
         TabIndex        =   17
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Instansi Penyelenggara"
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
         TabIndex        =   16
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Penyelenggaraan"
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
         Left            =   3720
         TabIndex        =   15
         Top             =   1800
         Width           =   1800
      End
   End
   Begin MSDataGridLib.DataGrid dgExtraPelatihan 
      Height          =   2895
      Left            =   120
      TabIndex        =   20
      Top             =   4560
      Width           =   10455
      _ExtentX        =   18441
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
      TabIndex        =   21
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
      Picture         =   "frmRiwayatExtraPelatihan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatExtraPelatihan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatExtraPelatihan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatExtraPelatihan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
    Call loaddc
    Call subLoadExtraPelatihan
    txtNamaPelatihan.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoUrut.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    strSQL = "DELETE FROM RiwayatExtraPelatihan WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL

    Call cmdBatal_Click

    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If dcjnspendidikan.Text <> "" Then
        If Periksa("datacombo", dcjnspendidikan, "Jenis Pendidikan Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If Periksa("text", txtNamaPelatihan, "Silahkan isi nama pelatihan ") = False Then Exit Sub
    If Periksa("text", txtKedudukanPeranan, "Silahkan isi kedudukan/peranan ") = False Then Exit Sub
    If Periksa("text", txtInstansiPenyelenggara, "Silahkan isi instansi penyelenggara ") = False Then Exit Sub

    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, Null)
        End If

        .Parameters.Append .CreateParameter("NamaPelatihan", adVarChar, adParamInput, 100, Trim(txtNamaPelatihan.Text))
        .Parameters.Append .CreateParameter("KedudukanPeranan", adVarChar, adParamInput, 50, Trim(txtKedudukanPeranan.Text))
        .Parameters.Append .CreateParameter("TglMulai", adDate, adParamInput, , Format(dtpTglMulai.Value, "yyyy/MM/dd"))
        If IsNull(dtpTglAkhir.Value) Then
            .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Format(dtpTglAkhir.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("LamaWaktu", adInteger, adParamInput, , IIf(txtlamawaktu.Text = "", Null, Trim(txtlamawaktu.Text)))
        .Parameters.Append .CreateParameter("InstansiPenyelenggara", adVarChar, adParamInput, 100, Trim(txtInstansiPenyelenggara.Text))
        .Parameters.Append .CreateParameter("AlamatPenyelenggaraan", adVarChar, adParamInput, 100, IIf(txtAlamatPenyelenggaraan.Text = "", Null, Trim(txtAlamatPenyelenggaraan.Text)))
        .Parameters.Append .CreateParameter("PimpinanPenyelenggara", adVarChar, adParamInput, 50, IIf(txtPimpinanPenyelenggara.Text = "", Null, Trim(txtPimpinanPenyelenggara.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("KdJenisPendidikan", adVarChar, adParamInput, 3, IIf(dcjnspendidikan.Text = "", Null, dcjnspendidikan.BoundText))
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 3, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RExtraPelatihan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Riwayat Extra Pelatihan Pegawai", vbCritical, "Validasi"
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
    Call subLoadExtraPelatihan
    Call subClearData
    Call loaddc
    txtNamaPelatihan.SetFocus
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcjnspendidikan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcjnspendidikan.Text)) = 0 Then txtNamaPelatihan.SetFocus: Exit Sub
        If dcjnspendidikan.MatchedWithList = True Then txtNamaPelatihan.SetFocus: Exit Sub
        strSQL = "select kdJenisPendidikan,jenispendidikan from jenispendidikan WHERE (JenisPendidikan LIKE '%" & dcjnspendidikan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcjnspendidikan.BoundText = rs(0).Value
        dcjnspendidikan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgExtraPelatihan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgExtraPelatihan.ApproxCount = 0 Then Exit Sub
    With dgExtraPelatihan
        txtNoUrut.Text = .Columns(1).Value
        If IsNull(.Columns(12)) Then dcjnspendidikan.BoundText = "" Else dcjnspendidikan.BoundText = .Columns(12).Value
        txtNamaPelatihan.Text = .Columns(2).Value
        txtKedudukanPeranan.Text = .Columns(3).Value
        dtpTglMulai.Value = .Columns(5).Value
        If IsNull(.Columns(6).Value) Then dtpTglAkhir.Value = Null Else dtpTglAkhir.Value = .Columns(6).Value
        If IsNull(.Columns(7).Value) Then txtlamawaktu.Text = "" Else txtlamawaktu.Text = .Columns(7).Value
        txtInstansiPenyelenggara.Text = .Columns(8).Value
        If IsNull(.Columns(9).Value) Then txtAlamatPenyelenggaraan.Text = "" Else txtAlamatPenyelenggaraan.Text = .Columns(9).Value
        If IsNull(.Columns(10).Value) Then txtPimpinanPenyelenggara.Text = "" Else txtPimpinanPenyelenggara.Text = .Columns(10).Value
        If IsNull(.Columns(11).Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = .Columns(11).Value
    End With
End Sub

Private Sub dtpTglAkhir_Change()
    dtpTglAkhir.MaxDate = Now
End Sub

Private Sub dtpTglAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtlamawaktu.SetFocus
End Sub

Private Sub dtpTglMulai_Change()
    dtpTglMulai.MaxDate = Now
End Sub

Private Sub dtpTglMulai_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then dtpTglAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadExtraPelatihan
    Call loaddc
End Sub

Private Sub subLoadExtraPelatihan()
    On Error GoTo hell

    strLSQL = "SELECT * FROM V_RiwayatExtraPelatihan WHERE IdPegawai='" & mstrIdPegawai & "'"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgExtraPelatihan
        Set .DataSource = rs
        .Columns(0).Width = 0           'IdPegawai
        .Columns(1).Width = 800
        .Columns(1).Caption = "No. Urut"
        .Columns(2).Width = 2000
        .Columns(2).Caption = " Nama Pelatihan"
        .Columns(3).Width = 2000
        .Columns(3).Caption = "Peranan"
        .Columns(4).Caption = "Jenis Pendidikan"
        .Columns(5).Width = 2100
        .Columns(5).Caption = "Tgl. Mulai"
        .Columns(6).Width = 1700
        .Columns(6).Caption = "Tgl. Akhir"
        .Columns(7).Width = 1700
        .Columns(7).Caption = "Lama Waktu"
        .Columns(8).Width = 2500
        .Columns(8).Caption = "Instansi Penyelenggara"
        .Columns(9).Width = 2500
        .Columns(9).Caption = "Alamat"
        .Columns(10).Caption = "Pimpinan"
        .Columns(11).Caption = "Keterangan"
        .Columns(12).Caption = "Nama User"

        .Columns(13).Width = 0
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    dcjnspendidikan.Text = ""
    txtNamaPelatihan.Text = ""
    txtKedudukanPeranan.Text = ""
    dtpTglMulai.Value = Format(Now, "dd/mmmm/yyyy")
    dtpTglAkhir.Value = Format(Now, "dd/mmmm/yyyy")
    txtlamawaktu.Text = ""
    txtInstansiPenyelenggara.Text = ""
    txtAlamatPenyelenggaraan.Text = ""
    txtPimpinanPenyelenggara.Text = ""
    txtKeterangan.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatExtraPelatihan
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtlamawaktu_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtInstansiPenyelenggara.SetFocus
End Sub

Private Sub txtNamaPelatihan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtKedudukanPeranan.SetFocus
End Sub

Private Sub txtKedudukanPeranan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglMulai.SetFocus
End Sub

Private Sub txtLamaPelatihan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtInstansiPenyelenggara.SetFocus
End Sub

Private Sub txtInstansiPenyelenggara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtAlamatPenyelenggaraan.SetFocus
End Sub

Private Sub txtAlamatPenyelenggaraan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtPimpinanPenyelenggara.SetFocus
End Sub

Private Sub txtPimpinanPenyelenggara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub loaddc()
    strSQL = "select kdJenisPendidikan,jenispendidikan from jenispendidikan where StatusEnabled='1' order by jenispendidikan "
    Call msubDcSource(dcjnspendidikan, rs, strSQL)
End Sub
