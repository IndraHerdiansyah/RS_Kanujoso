VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatPendidikanNonFormal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Pendidikan Non Formal Pegawai"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frmRiwayatPendidikanNonFormal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   9750
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
      Left            =   6840
      TabIndex        =   14
      Top             =   8160
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
      Left            =   5400
      TabIndex        =   13
      Top             =   8160
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
      Left            =   8280
      TabIndex        =   16
      Top             =   8160
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
      Left            =   3960
      TabIndex        =   12
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Pendidikan Non Formal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   9495
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
         MaxLength       =   100
         TabIndex        =   31
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txtPimpinanPendidikan 
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
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtTandaTanganSertifikat 
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
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtPendidikan 
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
         Left            =   3960
         MaxLength       =   100
         TabIndex        =   2
         Top             =   600
         Width           =   5295
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
         TabIndex        =   11
         Top             =   3480
         Width           =   9015
      End
      Begin VB.TextBox txtAlamatPendidikan 
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
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   9
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtNoSertifikat 
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
         Width           =   2535
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtLamaPendidikan 
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
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpTglMulai 
         Height          =   330
         Left            =   2880
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
         CustomFormat    =   "yyyy"
         Format          =   121110528
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpTglLulus 
         Height          =   330
         Left            =   5160
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
         CustomFormat    =   "yyyy"
         Format          =   121110528
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpTglSertifikat 
         Height          =   330
         Left            =   2880
         TabIndex        =   7
         Top             =   2040
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
         Format          =   121110528
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin MSDataListLib.DataCombo dcJenisPendidikan 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.Label Label8 
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
         TabIndex        =   32
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Sertifikat"
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
         TabIndex        =   30
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1440
         TabIndex        =   29
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pimpinan Pendidikan"
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
         TabIndex        =   28
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TandaTangan Sertifikat"
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
         TabIndex        =   27
         Top             =   1800
         Width           =   1680
      End
      Begin VB.Label Label30 
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
         TabIndex        =   24
         Top             =   3240
         Width           =   840
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Alamat Tempat Pendidikan"
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
         TabIndex        =   23
         Top             =   2520
         Width           =   1890
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "No Sertifikat"
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
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pendidikan Non Formal"
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
         Left            =   3960
         TabIndex        =   21
         Top             =   360
         Width           =   2070
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Lama Pendidikan"
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
         TabIndex        =   20
         Top             =   1080
         Width           =   1185
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
      Begin VB.Label Label10 
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
         Left            =   2880
         TabIndex        =   18
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Lulus"
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
   End
   Begin MSDataGridLib.DataGrid dgPendidikanNonFormal 
      Height          =   2775
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   9495
      _ExtentX        =   16748
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
      TabIndex        =   26
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
      Left            =   7920
      Picture         =   "frmRiwayatPendidikanNonFormal.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatPendidikanNonFormal.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPendidikanNonFormal.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatPendidikanNonFormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadPendidikanNonFormal
    dcJenisPendidikan.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoUrut.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPendidikanNonFormal WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If dcJenisPendidikan.Text <> "" Then
        If Periksa("datacombo", dcJenisPendidikan, "Jenis Pendidikan Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If Periksa("text", txtPendidikan, "Silahkan isi nama pendidikan ") = False Then Exit Sub
    If Periksa("text", txtLamaPendidikan, "Silahkan isi lama pendidikan ") = False Then Exit Sub
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
        .Parameters.Append .CreateParameter("NamaPendidikan", adVarChar, adParamInput, 100, Trim(txtPendidikan.Text))
        .Parameters.Append .CreateParameter("LamaPendidikan", adVarChar, adParamInput, 20, Trim(txtLamaPendidikan.Text))
        .Parameters.Append .CreateParameter("TglMulai", adDate, adParamInput, , Format(dtpTglMulai.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("TglLulus", adDate, adParamInput, , Format(dtpTglLulus.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("NoSertifikat", adVarChar, adParamInput, 50, IIf(txtNoSertifikat.Text = "", Null, Trim(txtNoSertifikat.Text)))
        If IsNull(dtpTglSertifikat.Value) Then
            .Parameters.Append .CreateParameter("TglSertifikat", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglSertifikat", adDate, adParamInput, , Format(dtpTglSertifikat.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("TandaTanganSertifikat", adVarChar, adParamInput, 30, IIf(txtTandaTanganSertifikat.Text = "", Null, Trim(txtTandaTanganSertifikat.Text)))
        .Parameters.Append .CreateParameter("InstansiPendidikan", adVarChar, adParamInput, 200, Trim(txtInstansiPenyelenggara.Text))
        .Parameters.Append .CreateParameter("AlamatPendidikan", adVarChar, adParamInput, 200, IIf(txtAlamatPendidikan.Text = "", Null, Trim(txtAlamatPendidikan.Text)))
        .Parameters.Append .CreateParameter("PimpinanPendidikan", adVarChar, adParamInput, 30, IIf(txtPimpinanPendidikan.Text = "", Null, Trim(txtPimpinanPendidikan.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("KdJenisPendidikan", adChar, adParamInput, 3, IIf(dcJenisPendidikan.Text = "", Null, dcJenisPendidikan.BoundText))
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 3, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RPddkNF"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Riwayat Pendidikan Non Formal pegawai", vbCritical, "Validasi"
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
    Call subLoadPendidikanNonFormal
    Call subClearData
    dcJenisPendidikan.SetFocus
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisPendidikan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtPendidikan.SetFocus

On Error GoTo Errload
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcJenisPendidikan.Text)) = 0 Then txtPendidikan.SetFocus: Exit Sub
        If dcJenisPendidikan.MatchedWithList = True Then txtPendidikan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "Select KdJenisPendidikan,JenisPendidikan from JenisPendidikan WHERE JenisPendidikan LIKE '%" & dcJenisPendidikan.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        dcJenisPendidikan.BoundText = dbRst(0).Value
        dcJenisPendidikan.Text = dbRst(1).Value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dgPendidikanNonFormal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgPendidikanNonFormal.ApproxCount = 0 Then Exit Sub
    txtNoUrut.Text = dgPendidikanNonFormal.Columns(1).Value
    txtPendidikan.Text = dgPendidikanNonFormal.Columns(2).Value
    txtLamaPendidikan.Text = dgPendidikanNonFormal.Columns(3).Value
    dtpTglMulai.Value = dgPendidikanNonFormal.Columns(4).Value
    dtpTglLulus.Value = dgPendidikanNonFormal.Columns(5).Value
    If IsNull(dgPendidikanNonFormal.Columns(6).Value) Then txtNoSertifikat.Text = "" Else txtNoSertifikat.Text = dgPendidikanNonFormal.Columns(6).Value
    If IsNull(dgPendidikanNonFormal.Columns(7).Value) Then dtpTglSertifikat.Value = Null Else dtpTglSertifikat.Value = dgPendidikanNonFormal.Columns(7).Value
    If IsNull(dgPendidikanNonFormal.Columns(8).Value) Then txtTandaTanganSertifikat.Text = "" Else txtTandaTanganSertifikat.Text = dgPendidikanNonFormal.Columns(8).Value
    If IsNull(dgPendidikanNonFormal.Columns(9).Value) Then txtInstansiPenyelenggara.Text = "" Else txtInstansiPenyelenggara.Text = dgPendidikanNonFormal.Columns(9).Value
    If IsNull(dgPendidikanNonFormal.Columns(10).Value) Then txtAlamatPendidikan.Text = "" Else txtAlamatPendidikan.Text = dgPendidikanNonFormal.Columns(10).Value
    If IsNull(dgPendidikanNonFormal.Columns(11).Value) Then txtPimpinanPendidikan.Text = "" Else txtPimpinanPendidikan.Text = dgPendidikanNonFormal.Columns(11).Value
    If IsNull(dgPendidikanNonFormal.Columns(12).Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgPendidikanNonFormal.Columns(12).Value
    If IsNull(dgPendidikanNonFormal.Columns(14).Value) Then dcJenisPendidikan.BoundText = "" Else dcJenisPendidikan.BoundText = dgPendidikanNonFormal.Columns(14).Value
End Sub

Private Sub dtpTglLulus_Change()
    dtpTglLulus.MaxDate = Now
End Sub

Private Sub dtpTglMulai_Change()
    dtpTglMulai.MaxDate = Now
End Sub

Private Sub dtpTglSertifikat_Change()
    dtpTglSertifikat.MaxDate = Now
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call SetComboJenisPendidikan
    Call subLoadPendidikanNonFormal
End Sub

Private Sub subLoadPendidikanNonFormal()
    On Error GoTo hell
    strLSQL = "SELECT * FROM RiwayatPendidikanNonFormal WHERE IdPegawai='" & mstrIdPegawai & "' ORDER BY NoUrut"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgPendidikanNonFormal
        Set .DataSource = rs
        .Columns(0).Width = 0           'IdPegawai
        .Columns(1).Width = 1000
        .Columns(1).Caption = "No. Urut"
        .Columns(2).Width = 2000
        .Columns(2).Caption = "Pendidikan"
        .Columns(3).Caption = "Lama Pendidikan"
        .Columns(3).Width = 1600
        .Columns(4).Caption = "Tgl. Mulai"
        .Columns(4).Width = 1500
        .Columns(5).Caption = "Tgl. Lulus"
        .Columns(5).Width = 1500
        .Columns(6).Caption = "No. Sertifikat"
        .Columns(7).Caption = "Tgl. Sertifikat"
        .Columns(8).Caption = "TTD Sertifikat"
        .Columns(9).Width = 3000
        .Columns(9).Caption = "Instansi Pendidikan"
        .Columns(10).Width = 3000
        .Columns(10).Caption = "Alamat Pendidikan"
        .Columns(11).Caption = "Pimpinan Pendidikan"
        .Columns(12).Width = 3500
        .Columns(13).Caption = "Nama User"
        .Columns(14).Width = 0
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    dcJenisPendidikan.Text = ""
    txtPendidikan.Text = ""
    txtLamaPendidikan.Text = ""
    dtpTglMulai.Value = Format(Now, "dd/mmmm/yyyy")
    dtpTglLulus.Value = Format(Now, "dd/mmmm/yyyy")
    txtNoSertifikat.Text = ""
    dtpTglSertifikat.Value = Format(Now, "dd/mmmm/yyyy")
    txtTandaTanganSertifikat.Text = ""
    txtInstansiPenyelenggara.Text = ""
    txtAlamatPendidikan.Text = ""
    txtPimpinanPendidikan.Text = ""
    txtKeterangan.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatPendidikanNonFormal
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtInstansiPenyelenggara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamatPendidikan.SetFocus
End Sub

Private Sub txtPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtLamaPendidikan.SetFocus
End Sub

Private Sub txtLamaPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglMulai.SetFocus
End Sub

Private Sub dtpTglMulai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglLulus.SetFocus
End Sub

Private Sub dtpTglLulus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNoSertifikat.SetFocus
End Sub

Private Sub txtNoSertifikat_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglSertifikat.SetFocus
End Sub

Private Sub dtpTglSertifikat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTandaTanganSertifikat.SetFocus
End Sub

Private Sub txtAlamatPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtPimpinanPendidikan.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtPimpinanPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtTandaTanganSertifikat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtInstansiPenyelenggara.SetFocus
End Sub

Sub SetComboJenisPendidikan()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from JenisPendidikan where KdJenisPendidikan='002'", dbConn, , adLockOptimistic
    Set dcJenisPendidikan.RowSource = rs
    dcJenisPendidikan.ListField = rs.Fields(1).Name
    dcJenisPendidikan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Exit Sub
hell:
    Call msubPesanError
End Sub
