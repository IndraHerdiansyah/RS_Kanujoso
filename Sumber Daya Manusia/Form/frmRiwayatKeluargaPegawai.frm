VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatKeluargaPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Keluarga Pegawai"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   Icon            =   "frmRiwayatKeluargaPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   9495
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
      Left            =   6600
      TabIndex        =   11
      Top             =   7440
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
      Left            =   5160
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
      Left            =   8040
      TabIndex        =   12
      Top             =   7440
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
      Left            =   3720
      TabIndex        =   9
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Keluarga Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   9255
      Begin VB.TextBox txtAlamat 
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
         Top             =   3000
         Width           =   8775
      End
      Begin VB.TextBox txtTempat 
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
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   27
         Top             =   2400
         Width           =   3495
      End
      Begin VB.ComboBox cbStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmRiwayatKeluargaPegawai.frx":0CCA
         Left            =   8160
         List            =   "frmRiwayatKeluargaPegawai.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1800
         Width           =   855
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
         Height          =   315
         Left            =   240
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1800
         Width           =   7815
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
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtNmLengkap 
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
         Height          =   315
         Left            =   3720
         MaxLength       =   30
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin MSDataListLib.DataCombo dcKdHubungan 
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
      Begin MSComCtl2.DTPicker dtpTglLahir 
         Height          =   330
         Left            =   240
         TabIndex        =   4
         Top             =   1200
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
         Format          =   99745792
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin MSDataListLib.DataCombo dcKdPekerjaan 
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo dcKdPendidikan 
         Height          =   315
         Left            =   6120
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
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
      Begin MSDataListLib.DataCombo dcJK 
         Height          =   315
         Left            =   7680
         TabIndex        =   3
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSComCtl2.DTPicker dtpTglKawin 
         Height          =   330
         Left            =   240
         TabIndex        =   25
         Top             =   2400
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
         Format          =   120979456
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSDataListLib.DataCombo dcAgama 
         Height          =   315
         Left            =   6120
         TabIndex        =   29
         Top             =   2400
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
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
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Agama"
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
         TabIndex        =   30
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tempat Nikah"
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
         TabIndex        =   28
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Nikah"
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
         TabIndex        =   26
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Lahir"
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
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan"
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
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Hubungan Keluarga"
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
         TabIndex        =   21
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pendidikan"
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
         TabIndex        =   20
         Top             =   960
         Width           =   765
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
         TabIndex        =   19
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Pekerjaan"
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
         TabIndex        =   18
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "JK"
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
         TabIndex        =   17
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nama Lengkap"
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
         TabIndex        =   16
         Top             =   360
         Width           =   1050
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
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dgKeluarga 
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Left            =   7680
      Picture         =   "frmRiwayatKeluargaPegawai.frx":0CDE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatKeluargaPegawai.frx":1A66
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatKeluargaPegawai.frx":4427
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatKeluargaPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cbStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dtpTglKawin.Enabled = False Then
            cmdSimpan.SetFocus
        Else
            dtpTglKawin.SetFocus
        End If
    End If
End Sub

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadKeluarga
    dcKdHubungan.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoUrut.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM KeluargaPegawai WHERE IdPegawai='" & mstrIdPegawai & "' AND KdHubungan = '" & dcKdHubungan.BoundText & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If dcKdHubungan.Text <> "" Then
        If Periksa("datacombo", dcKdHubungan, "Hubungan Keluarga Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcJK.Text <> "" Then
        If Periksa("datacombo", dcJK, "Jenis Kelamin Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcKdPekerjaan.Text <> "" Then
        If Periksa("datacombo", dcKdPekerjaan, "Pekerjaan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcKdPendidikan.Text <> "" Then
        If Periksa("datacombo", dcKdPendidikan, "Pendidikan Tidak Terdaftar") = False Then Exit Sub
    End If
    If dcAgama.Text <> "" Then
        If Periksa("datacombo", dcAgama, "Agama Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If Periksa("dataCombo", dcKdHubungan, "Silahkan isi hubungan keluarga ") = False Then Exit Sub
    If Periksa("text", txtNmLengkap, "Silahkan isi nama lengkap ") = False Then Exit Sub
    If Periksa("datacombo", dcJK, "Silahkan isi jenis kelamin ") = False Then Exit Sub

    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        .Parameters.Append .CreateParameter("KdHubungan", adChar, adParamInput, 2, dcKdHubungan.BoundText)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Null)
        End If
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 30, Trim(txtNmLengkap.Text))
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, dcJK.BoundText)
        If IsNull(dtpTglLahir.Value) Then
            .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(dtpTglLahir.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("KdPekerjaan", adChar, adParamInput, 2, IIf(dcKdPekerjaan.Text = "", Null, dcKdPekerjaan.BoundText))
        .Parameters.Append .CreateParameter("KdPendidikan", adChar, adParamInput, 4, IIf(dcKdPendidikan.Text = "", Null, dcKdPendidikan.BoundText))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("StatusTanggungan", adChar, adParamInput, 1, cbStatus.Text)
        If IsNull(dtpTglKawin.Value) Then
            .Parameters.Append .CreateParameter("TglNikah", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglNikah", adDate, adParamInput, , Format(dtpTglKawin.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("TempatNikah", adVarChar, adParamInput, 50, IIf(txtTempat.Text = "", Null, Trim(txtTempat.Text)))
        .Parameters.Append .CreateParameter("KdAgama", adChar, adParamInput, 2, IIf(dcAgama.Text = "", Null, dcAgama.BoundText))
        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, IIf(txtAlamat.Text = "", Null, Trim(txtAlamat.Text)))
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 2, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_KeluargaPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Keluarga pegawai", vbCritical, "Validasi"
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
    Call subLoadKeluarga
    Call subClearData
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAgama_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtAlamat.SetFocus

On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcAgama.Text)) = 0 Then txtKeterangan.SetFocus: Exit Sub
        If dcAgama.MatchedWithList = True Then txtAlamat.SetFocus: Exit Sub
        strSQL = "Select KdAgama, Agama from Agama WHERE (Agama LIKE '%" & dcAgama.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcAgama.BoundText = rs(0).Value
        dcAgama.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcJK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglLahir.SetFocus
End Sub

Private Sub dcJK_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcJK.Text)) = 0 Then dtpTglLahir.SetFocus: Exit Sub
        If dcJK.MatchedWithList = True Then dtpTglLahir.SetFocus: Exit Sub
        strSQL = "SELECT KdJenisKelamin, JenisKelamin FROM JenisKelamin WHERE (JenisKelamin LIKE '%" & dcJK.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcJK.BoundText = rs(0).Value
        dcJK.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKdHubungan_Change()
    If dcKdHubungan.BoundText = "01" Then
        dtpTglKawin.Enabled = True
        txtTempat.Enabled = True
        dcAgama.Enabled = True
        txtAlamat.Enabled = True

    ElseIf dcKdHubungan.BoundText = "02" Then
        dtpTglKawin.Enabled = True
        txtTempat.Enabled = True
        dcAgama.Enabled = True
        txtAlamat.Enabled = True

    Else
        dtpTglKawin.Enabled = False
        txtTempat.Enabled = False
        dcAgama.Enabled = False
        txtAlamat.Enabled = False

    End If
End Sub

Private Sub dcKdHubungan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKdHubungan.Text)) = 0 Then txtNmLengkap.SetFocus: Exit Sub
        If dcKdHubungan.MatchedWithList = True Then txtNmLengkap.SetFocus: Exit Sub
        strSQL = "Select Hubungan,NamaHubungan from HubunganKeluarga WHERE (NamaHubungan LIKE '%" & dcKdHubungan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKdHubungan.BoundText = rs(0).Value
        dcKdHubungan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKdPekerjaan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKdPekerjaan.Text)) = 0 Then dcKdPendidikan.SetFocus: Exit Sub
        If dcKdPekerjaan.MatchedWithList = True Then dcKdPendidikan.SetFocus: Exit Sub
        strSQL = "Select KdPekerjaan,Pekerjaan from Pekerjaan WHERE (Pekerjaan LIKE '%" & dcKdPekerjaan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKdPekerjaan.BoundText = rs(0).Value
        dcKdPekerjaan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKdPendidikan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKdPendidikan.Text)) = 0 Then txtKeterangan.SetFocus: Exit Sub
        If dcKdPendidikan.MatchedWithList = True Then txtKeterangan.SetFocus: Exit Sub
        strSQL = "Select KdPendidikan,Pendidikan from Pendidikan WHERE (Pendidikan LIKE '%" & dcKdPendidikan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKdPendidikan.BoundText = rs(0).Value
        dcKdPendidikan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgKeluarga_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtNoUrut.Text = dgKeluarga.Columns("NoUrut").Value
    dcKdHubungan.BoundText = dgKeluarga.Columns("KdHubungan").Value
    txtNmLengkap.Text = dgKeluarga.Columns("NamaLengkap").Value
    dcJK.Text = dgKeluarga.Columns("JenisKelamin")
    If dgKeluarga.Columns("JenisKelamin").Value = "L" Then dcJK.Text = "Laki-Laki" Else dcJK.Text = "Perempuan"
    If IsNull(dgKeluarga.Columns("TglLahir").Value) Then dtpTglLahir.Value = Null Else dtpTglLahir.Value = dgKeluarga.Columns("TglLahir").Value
    If IsNull(dgKeluarga.Columns("KdPekerjaan").Value) Then dcKdPekerjaan.BoundText = "" Else dcKdPekerjaan.BoundText = dgKeluarga.Columns("KdPekerjaan").Value
    If IsNull(dgKeluarga.Columns("KdPendidikan").Value) Then dcKdPendidikan.BoundText = "" Else dcKdPendidikan.Text = dgKeluarga.Columns("KdPendidikan").Value
    If IsNull(dgKeluarga.Columns("Keterangan").Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgKeluarga.Columns("Keterangan").Value
    If IsNull(dgKeluarga.Columns("StatusTanggungan")) Then cbStatus.Text = "" Else cbStatus.Text = dgKeluarga.Columns("StatusTanggungan")
    If IsNull(dgKeluarga.Columns("TglNikah").Value) Then dtpTglKawin.Value = Null Else dtpTglKawin.Value = dgKeluarga.Columns("TglNikah").Value
    If IsNull(dgKeluarga.Columns("TempatNikah").Value) Then txtTempat.Text = "" Else txtTempat.Text = dgKeluarga.Columns("TempatNikah").Value
    If IsNull(dgKeluarga.Columns("KdAgama").Value) Then dcAgama.BoundText = "" Else dcAgama.BoundText = dgKeluarga.Columns("KdAgama").Value
    If IsNull(dgKeluarga.Columns("Alamat").Value) Then txtAlamat.Text = "" Else txtAlamat.Text = dgKeluarga.Columns("Alamat").Value

End Sub

Private Sub dtpTglKawin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTempat.SetFocus
End Sub

Private Sub dtpTglLahir_Change()
    dtpTglLahir.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call SetComboHubungan
    Call SetComboJK
    Call SetComboPekerjaan
    Call SetComboPendidikan
    Call SetComboAgama
    Call subLoadKeluarga
End Sub

Private Sub subLoadKeluarga()
    On Error GoTo hell
    strLSQL = "SELECT * FROM V_KeluargaPegawai WHERE IdPegawai='" & mstrIdPegawai & "' ORDER BY NoUrut"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgKeluarga
        Set .DataSource = rs
        .Columns(0).Width = 0           'IdPegawai
        .Columns(1).Width = 0
        .Columns(2).Width = 1000
        .Columns(3).Width = 1500
        .Columns(4).Width = 2000
        .Columns(5).Width = 1000
        .Columns(6).Width = 1400
        .Columns(7).Width = 1400
        .Columns(8).Width = 1400
        .Columns(9).Width = 2000
        .Columns(10).Width = 1500
        .Columns(11).Width = 1500
        .Columns(12).Width = 1500
        .Columns(13).Width = 1500
        .Columns(14).Width = 2000
        .Columns(15).Width = 0
        .Columns(16).Width = 0
        .Columns(17).Width = 0
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    dcKdHubungan.Text = ""
    txtNmLengkap.Text = ""
    dtpTglLahir.Value = Format(Now, "dd/mm/yyyy")
    dcJK.Text = ""
    dcKdPekerjaan.Text = ""
    dcKdPendidikan.Text = ""
    txtKeterangan.Text = ""
    dtpTglKawin.Value = Format(Now, "dd/mm/yyyy")
    txtTempat.Text = ""
    dcAgama.BoundText = ""
    txtAlamat.Text = ""
End Sub

Sub SetComboHubungan()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from HubunganKeluarga where StatusEnabled='1'", dbConn, , adLockOptimistic
    Set dcKdHubungan.RowSource = rs
    dcKdHubungan.ListField = rs.Fields(1).Name
    dcKdHubungan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub SetComboPekerjaan()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from Pekerjaan where StatusEnabled='1'", dbConn, , adLockOptimistic
    Set dcKdPekerjaan.RowSource = rs
    dcKdPekerjaan.ListField = rs.Fields(1).Name
    dcKdPekerjaan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub SetComboPendidikan()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from Pendidikan where StatusEnabled='1'", dbConn, , adLockOptimistic
    Set dcKdPendidikan.RowSource = rs
    dcKdPendidikan.ListField = rs.Fields(1).Name
    dcKdPendidikan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub SetComboJK()
    Set rs = Nothing
    strSQL = "SELECT Singkatan, JenisKelamin FROM JenisKelamin"
    Call msubDcSource(dcJK, rs, strSQL)
    Set rs = Nothing
End Sub

Sub SetComboAgama()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select KdAgama, Agama from Agama where StatusEnabled='1'", dbConn, , adLockOptimistic
    Set dcAgama.RowSource = rs
    dcAgama.ListField = rs.Fields(1).Name
    dcAgama.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcKdHubungan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If dcKdHubungan.BoundText = "01" Then
            dtpTglKawin.Enabled = True
            txtTempat.Enabled = True
            dcAgama.Enabled = True
            txtAlamat.Enabled = True
            txtNmLengkap.SetFocus
        ElseIf dcKdHubungan.BoundText = "02" Then
            dtpTglKawin.Enabled = True
            txtTempat.Enabled = True
            dcAgama.Enabled = True
            txtAlamat.Enabled = True
            txtNmLengkap.SetFocus
        Else
            dtpTglKawin.Enabled = False
            txtTempat.Enabled = False
            dcAgama.Enabled = False
            txtAlamat.Enabled = False
            txtNmLengkap.SetFocus
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatKeluargaPegawai
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNmLengkap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dcJK.SetFocus
End Sub

Private Sub cbJK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglLahir.SetFocus
End Sub

Private Sub dtpTglLahir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKdPekerjaan.SetFocus
End Sub

Private Sub dcKdPekerjaan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKdPendidikan.SetFocus
End Sub

Private Sub dcKdPendidikan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cbStatus.SetFocus
End Sub

Private Sub txtTempat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcAgama.SetFocus
End Sub
