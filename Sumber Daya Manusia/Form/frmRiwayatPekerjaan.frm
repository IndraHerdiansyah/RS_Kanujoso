VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatPekerjaan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Pekerjaan Pegawai"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmRiwayatPekerjaan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10230
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
      Left            =   4440
      TabIndex        =   12
      Top             =   8040
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
      Left            =   8760
      TabIndex        =   15
      Top             =   8040
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
      Left            =   5880
      TabIndex        =   13
      Top             =   8040
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
      Left            =   7320
      TabIndex        =   14
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Pekerjaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   9975
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtPimpinanPerusahaan 
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
         Left            =   5640
         MaxLength       =   30
         TabIndex        =   9
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtGajiPokok 
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
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtUraianPekerjaan 
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
         TabIndex        =   11
         Top             =   3480
         Width           =   9375
      End
      Begin VB.TextBox txtNamaPerusahaan 
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
         Width           =   4335
      End
      Begin VB.TextBox txtJabatanPosisi 
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
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3855
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
         MaxLength       =   2
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtAlamatPerusahaan 
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
         TabIndex        =   10
         Top             =   2760
         Width           =   9375
      End
      Begin VB.TextBox txtNoSuratKeputusan 
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
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1320
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtpTglMulai 
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   16449536
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   330
         Left            =   2400
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
         Format          =   16449536
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpTglSK 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   2040
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
         Format          =   16449536
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   435
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
         Left            =   2400
         TabIndex        =   29
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
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
         Left            =   2640
         TabIndex        =   28
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pimpinan Perusahaan"
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
         Left            =   5640
         TabIndex        =   27
         Top             =   1800
         Width           =   1530
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Uraian Pekerjaan"
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
         TabIndex        =   25
         Top             =   3240
         Width           =   1230
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Gaji Pokok"
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
         Left            =   4800
         TabIndex        =   24
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label12 
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
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan/Posisi"
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
         Left            =   5880
         TabIndex        =   20
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Perusahaan"
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
         TabIndex        =   19
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Perusahaan"
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
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "No Surat Keputusan"
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
         TabIndex        =   17
         Top             =   1080
         Width           =   1440
      End
   End
   Begin MSDataGridLib.DataGrid dgPekerjaan 
      Height          =   2655
      Left            =   120
      TabIndex        =   23
      Top             =   5280
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4683
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
      TabIndex        =   26
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8400
      Picture         =   "frmRiwayatPekerjaan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatPekerjaan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPekerjaan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatPekerjaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadPekerjaan
    txtNamaPerusahaan.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoUrut.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPekerjaan WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If Periksa("text", txtNamaPerusahaan, "Silahkan isi nama perusahaan ") = False Then Exit Sub
    If Periksa("text", txtJabatanPosisi, "Silahkan nama posisi/jabatan ") = False Then Exit Sub
    If Periksa("text", txtGajiPokok, "Gaji pokok silahkan diisi ") = False Then Exit Sub
    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Null)
        End If
        .Parameters.Append .CreateParameter("NamaPerusahaan", adVarChar, adParamInput, 100, Trim(txtNamaPerusahaan.Text))
        .Parameters.Append .CreateParameter("JabatanPosisi", adVarChar, adParamInput, 50, Trim(txtJabatanPosisi.Text))
        .Parameters.Append .CreateParameter("UraianPekerjaan", adVarChar, adParamInput, 200, IIf(txtUraianPekerjaan.Text = "", Null, Trim(txtUraianPekerjaan.Text)))
        .Parameters.Append .CreateParameter("TglMulai", adDate, adParamInput, , Format(dtpTglMulai.Value, "yyyy/MM/dd"))
        If IsNull(dtpTglAkhir.Value) Then
            .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Format(dtpTglAkhir.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("GajiPokok", adInteger, adParamInput, , txtGajiPokok.Text)
        .Parameters.Append .CreateParameter("NoSK", adVarChar, adParamInput, 50, IIf(txtNoSuratKeputusan.Text = "", Null, Trim(txtNoSuratKeputusan.Text)))
        If IsNull(dtpTglSK.Value) Then
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Format(dtpTglSK.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("TandaTanganSK", adVarChar, adParamInput, 30, IIf(txtTandaTanganSK.Text = "", Null, Trim(txtTandaTanganSK.Text)))
        .Parameters.Append .CreateParameter("AlamatPerusahaan", adVarChar, adParamInput, 200, IIf(txtAlamatPerusahaan.Text = "", Null, Trim(txtAlamatPerusahaan.Text)))
        .Parameters.Append .CreateParameter("PimpinanPerusahaan", adVarChar, adParamInput, 30, IIf(txtPimpinanPerusahaan.Text = "", Null, Trim(txtPimpinanPerusahaan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 2, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RPekerjaan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Riwayat Pekerjaan pegawai", vbCritical, "Validasi"
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
    Call subLoadPekerjaan
    Call subClearData
    txtNamaPerusahaan.SetFocus
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgPekerjaan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgPekerjaan.ApproxCount = 0 Then Exit Sub
    txtNoUrut.Text = dgPekerjaan.Columns("NoUrut").Value
    txtNamaPerusahaan.Text = dgPekerjaan.Columns("NamaPerusahaan").Value
    txtJabatanPosisi.Text = dgPekerjaan.Columns("JabatanPosisi").Value
    If IsNull(dgPekerjaan.Columns("UraianPekerjaan").Value) Then txtUraianPekerjaan.Text = "" Else txtUraianPekerjaan.Text = dgPekerjaan.Columns("UraianPekerjaan").Value
    dtpTglMulai.Value = dgPekerjaan.Columns("TglMulai").Value
    If IsNull(dgPekerjaan.Columns("TglAkhir").Value) Then dtpTglAkhir.Value = Null Else dtpTglAkhir.Value = dgPekerjaan.Columns("TglAkhir").Value
    txtGajiPokok.Text = dgPekerjaan.Columns("GajiPokok").Value
    Call txtGajiPokok_LostFocus
    If IsNull(dgPekerjaan.Columns("NoSK").Value) Then txtNoSuratKeputusan.Text = "" Else txtNoSuratKeputusan.Text = dgPekerjaan.Columns("NoSK").Value
    If IsNull(dgPekerjaan.Columns("TglSK").Value) Then dtpTglSK.Value = Null Else dtpTglSK.Value = dgPekerjaan.Columns("TglSK").Value
    If IsNull(dgPekerjaan.Columns("TandaTanganSK").Value) Then txtTandaTanganSK.Text = "" Else txtTandaTanganSK.Text = dgPekerjaan.Columns("TandaTanganSK").Value
    If IsNull(dgPekerjaan.Columns("AlamatPerusahaan").Value) Then txtAlamatPerusahaan.Text = "" Else txtAlamatPerusahaan.Text = dgPekerjaan.Columns("AlamatPerusahaan").Value
    If IsNull(dgPekerjaan.Columns("PimpinanPerusahaan").Value) Then txtPimpinanPerusahaan.Text = "" Else txtPimpinanPerusahaan.Text = dgPekerjaan.Columns("PimpinanPerusahaan").Value
End Sub

Private Sub dtpTglAkhir_Change()
'    dtpTglAkhir.MaxDate = Now
End Sub

Private Sub dtpTglMulai_Change()
    dtpTglMulai.MaxDate = Now
End Sub

Private Sub dtpTglSK_Change()
    dtpTglSK.MaxDate = Now
End Sub

Private Sub dtpTglSK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTandaTanganSK.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadPekerjaan
End Sub

Private Sub subLoadPekerjaan()
    On Error GoTo hell
    strLSQL = "SELECT * FROM RiwayatPekerjaan WHERE IdPegawai='" & mstrIdPegawai & "' ORDER BY NoUrut"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgPekerjaan
        Set .DataSource = rs
        .Columns("IdPegawai").Width = 0           'IdPegawai
        .Columns("NoUrut").Width = 1000
        .Columns("NamaPerusahaan").Width = 2000
        .Columns("JabatanPosisi").Width = 1600
        .Columns("UraianPekerjaan").Width = 1800
        .Columns("TglMulai").Width = 1200
        .Columns("TglAkhir").Width = 1200
        .Columns("GajiPokok").Width = 2000
        .Columns("GajiPokok").NumberFormat = "#,###"
        .Columns("NoSK").Width = 1500
        .Columns("TglSK").Width = 1200
        .Columns("TandaTanganSK").Width = 1800
        .Columns("AlamatPerusahaan").Width = 2200
        .Columns("PimpinanPerusahaan").Width = 2000
        .Columns("IdUser").Width = 0
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    txtNamaPerusahaan.Text = ""
    txtJabatanPosisi.Text = ""
    txtUraianPekerjaan.Text = ""
    dtpTglMulai.Value = Format(Now, "dd/mmmm/yyyy")
    dtpTglAkhir.Value = Format(Now, "dd/mmmm/yyyy")
    txtGajiPokok.Text = 0
    txtNoSuratKeputusan.Text = ""
    dtpTglSK.Value = Format(Now, "dd/mmmm/yyyy")
    txtTandaTanganSK.Text = ""
    txtAlamatPerusahaan.Text = ""
    txtPimpinanPerusahaan.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatPekerjaan
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtGajiPokok_LostFocus()
    txtGajiPokok.Text = IIf(Val(txtGajiPokok) = 0, 0, Format(txtGajiPokok, "#,###"))
End Sub

Private Sub txtNamaPerusahaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtJabatanPosisi.SetFocus
End Sub

Private Sub txtJabatanPosisi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglMulai.SetFocus
End Sub

Private Sub txtPimpinanPerusahaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamatPerusahaan.SetFocus
End Sub

Private Sub txtTandaTanganSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPimpinanPerusahaan.SetFocus
End Sub

Private Sub txtUraianPekerjaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dtpTglMulai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglAkhir.SetFocus
End Sub

Private Sub dtpTglAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtGajiPokok.SetFocus
End Sub

Private Sub txtGajiPokok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNoSuratKeputusan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNoSuratKeputusan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglSK.SetFocus
End Sub

Private Sub txtAlamatPerusahaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtUraianPekerjaan.SetFocus
End Sub
