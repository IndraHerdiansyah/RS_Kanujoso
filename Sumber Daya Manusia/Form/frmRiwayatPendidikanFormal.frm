VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatPendidikanFormal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Pendidikan Formal Pegawai"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   Icon            =   "frmRiwayatPendidikanFormal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10935
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
      TabIndex        =   14
      Top             =   7800
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
      TabIndex        =   17
      Top             =   7800
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
      TabIndex        =   15
      Top             =   7800
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
      Left            =   8040
      TabIndex        =   16
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Pendidikan Formal"
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
      TabIndex        =   20
      Top             =   1080
      Width           =   10695
      Begin VB.TextBox txtKodeJenisPendidikan 
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtIPK 
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
         ForeColor       =   &H80000006&
         Height          =   330
         Left            =   9720
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1320
         Width           =   735
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
         Left            =   240
         MaxLength       =   100
         TabIndex        =   11
         Top             =   2760
         Width           =   6855
      End
      Begin VB.TextBox txtNoIjazah 
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
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2040
         Width           =   2415
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
         Width           =   615
      End
      Begin VB.TextBox txtFakultasJurusan 
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
         Width           =   4575
      End
      Begin VB.ComboBox cbGradeKelulusan 
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
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtNamaPendidikan 
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
         Left            =   3360
         MaxLength       =   100
         TabIndex        =   2
         Top             =   600
         Width           =   7095
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
         Left            =   7200
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox txtTandaTanganIjazah 
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
         Left            =   7200
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2040
         Width           =   3255
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
         TabIndex        =   13
         Top             =   3480
         Width           =   10215
      End
      Begin MSComCtl2.DTPicker dtpTglLulus 
         Height          =   330
         Left            =   7200
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   120193027
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpTglIjazah 
         Height          =   330
         Left            =   4920
         TabIndex        =   9
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
         Format          =   120193024
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin MSDataListLib.DataCombo dcPendidikan 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   600
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
      Begin MSComCtl2.DTPicker dtpTglMasuk 
         Height          =   330
         Left            =   4920
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
         Format          =   120193024
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Ijazah"
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
         TabIndex        =   34
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label Label4 
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
         Left            =   4920
         TabIndex        =   33
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label3 
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
         Left            =   7200
         TabIndex        =   32
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label30 
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
         Left            =   7200
         TabIndex        =   31
         Top             =   2520
         Width           =   1440
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
         Left            =   240
         TabIndex        =   30
         Top             =   2520
         Width           =   1890
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "No. Ijazah"
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
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Tingkat Kelulusan"
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
         TabIndex        =   28
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label7 
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
         Left            =   960
         TabIndex        =   27
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nilai IPK"
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
         Left            =   9720
         TabIndex        =   26
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jurusan"
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
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No."
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
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Tempat Pendidikan"
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
         Left            =   3360
         TabIndex        =   23
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tanda Tangan Ijazah"
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
         TabIndex        =   22
         Top             =   1800
         Width           =   1530
      End
      Begin VB.Label Label13 
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
         TabIndex        =   21
         Top             =   3240
         Width           =   840
      End
   End
   Begin MSDataGridLib.DataGrid dgPendidikanFormal 
      Height          =   2535
      Left            =   120
      TabIndex        =   19
      Top             =   5160
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Left            =   9120
      Picture         =   "frmRiwayatPendidikanFormal.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatPendidikanFormal.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPendidikanFormal.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatPendidikanFormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub chkTglLulus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtIPK.SetFocus
End Sub

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadPendidikanFormal
    cmdSimpan.Enabled = True
    dcpendidikan.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoUrut.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPendidikanFormal WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL
    MsgBox "Data berhasil diihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If dcpendidikan.Text <> "" Then
        If Periksa("datacombo", dcpendidikan, "Pendidikan Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If Periksa("datacombo", dcpendidikan, "Silahkan isi nama pendidikannya ") = False Then Exit Sub
    If Periksa("text", txtNamaPendidikan, "Silahkan isi nama tempat pendidikannya ") = False Then Exit Sub

    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        .Parameters.Append .CreateParameter("KdPendidikan", adChar, adParamInput, 4, dcpendidikan.BoundText)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Null)
        End If
        .Parameters.Append .CreateParameter("NamaPendidikan", adVarChar, adParamInput, 100, Trim(txtNamaPendidikan.Text))
        .Parameters.Append .CreateParameter("FakultasJurusan", adVarChar, adParamInput, 100, IIf(txtFakultasJurusan.Text = "", Null, Trim(txtFakultasJurusan.Text)))
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglMasuk.Value, "yyyy/MM/dd"))
        If IsNull(dtpTglLulus.Value) Then
            .Parameters.Append .CreateParameter("TglLulus", adDate, adParamInput, , Null)
        Else

'            strSQL = "Select TglLulus from RiwayatPendidikanFormal where IdPegawai = '" & mstrIdPegawai & "' And TglLulus = '" & Format(dtpTglLulus.Value, "yyyy/MM/dd") & "'"
'            Call msubRecFO(rs, strSQL)
'            If rs.EOF = False Or rs.BOF = False Then
'                Call MsgBox("Ada Tanggal lulus yang sama, Cek kembali tanggal kelulusan", vbOKOnly, "Medifirst2000-Validasi")
'                Call deleteADOCommandParameters(adoCommand)
'                Exit Sub
'            End If
            .Parameters.Append .CreateParameter("TglLulus", adDate, adParamInput, , Format(dtpTglLulus.Value, "yyyy/MM/dd"))

        End If
        If txtIPK.Text = "" Then
            .Parameters.Append .CreateParameter("IPK", adDouble, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("IPK", adDouble, adParamInput, , CDec(txtIPK.Text))
        End If
        .Parameters.Append .CreateParameter("GradeKelulusan", adVarChar, adParamInput, 50, IIf(cbGradeKelulusan.Text = "", Null, Trim(cbGradeKelulusan.Text)))
        .Parameters.Append .CreateParameter("NoIjazah", adVarChar, adParamInput, 30, IIf(txtNoIjazah.Text = "", Null, Trim(txtNoIjazah.Text)))
        If IsNull(dtpTglIjazah.Value) Then
            .Parameters.Append .CreateParameter("TglIjazah", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglIjazah", adDate, adParamInput, , Format(dtpTglIjazah.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("TandaTanganIjazah", adVarChar, adParamInput, 30, IIf(txtTandaTanganIjazah.Text = "", Null, Trim(txtTandaTanganIjazah.Text)))
        .Parameters.Append .CreateParameter("AlamatPendidikan", adVarChar, adParamInput, 200, IIf(txtAlamatPendidikan.Text = "", Null, Trim(txtAlamatPendidikan.Text)))
        .Parameters.Append .CreateParameter("PimpinanPendidikan", adChar, adParamInput, 30, IIf(txtPimpinanPendidikan.Text = "", Null, Trim(txtPimpinanPendidikan.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("KdJenisPendidikan", adVarChar, adParamInput, 3, IIf(txtKodeJenisPendidikan.Text = "", Null, Trim(txtKodeJenisPendidikan.Text)))
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 2, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RPddkF"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Riwayat Pendidikan Formal Pegawai", vbCritical, "Validasi"
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
    Call subLoadPendidikanFormal
    Call subClearData
    dcpendidikan.SetFocus
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Call frmRiwayatPegawai.subLoadRiwayatPendidikanFormal
    frmRiwayatPegawai.Enabled = True
    Unload Me
End Sub

Private Sub dcPendidikan_Change()
    On Error Resume Next
    strSQL = "select KdJenisPendidikan from Pendidikan where KdPendidikan='" & dcpendidikan.BoundText & "'"
    Call msubRecFO(rsSplakuk, strSQL)
    If rsSplakuk.EOF = True Then txtKodeJenisPendidikan.Text = "" Else txtKodeJenisPendidikan.Text = rsSplakuk(0).Value
End Sub

Private Sub dcPendidikan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtNamaPendidikan.SetFocus

On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcpendidikan.Text)) = 0 Then txtNamaPendidikan.SetFocus: Exit Sub
        If dcpendidikan.MatchedWithList = True Then txtNamaPendidikan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdPendidikan, Pendidikan FROM Pendidikan WHERE Pendidikan LIKE '%" & dcpendidikan.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        dcpendidikan.BoundText = dbRst(0).Value
        dcpendidikan.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPendidikan_LostFocus()
    On Error Resume Next
    strSQL = "select KdJenisPendidikan from Pendidikan where KdPendidikan='" & dcpendidikan.BoundText & "'"
    Call msubRecFO(rsSplakuk, strSQL)
    If rsSplakuk.EOF = True Then txtKodeJenisPendidikan.Text = "" Else txtKodeJenisPendidikan.Text = rsSplakuk(0).Value
End Sub

Private Sub dgPendidikanFormal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgPendidikanFormal.ApproxCount = 0 Then Exit Sub
    txtNoUrut.Text = dgPendidikanFormal.Columns("No. Urut").Value
    dcpendidikan.Text = dgPendidikanFormal.Columns("Pendidikan").Value
    txtNamaPendidikan.Text = dgPendidikanFormal.Columns("Nama Sekolah").Value
    If IsNull(dgPendidikanFormal.Columns("Jurusan").Value) Then txtFakultasJurusan.Text = "" Else txtFakultasJurusan.Text = dgPendidikanFormal.Columns("Jurusan").Value
    dtpTglMasuk.Value = dgPendidikanFormal.Columns("Tgl. Masuk").Value
    If IsNull(dgPendidikanFormal.Columns("Tgl. Lulus").Value) Then dtpTglLulus.Value = Null Else dtpTglLulus.Value = dgPendidikanFormal.Columns("Tgl. Lulus").Value
    If IsNull(dgPendidikanFormal.Columns("IPK").Value) Then txtIPK.Text = "" Else txtIPK.Text = dgPendidikanFormal.Columns("IPK").Value
    If IsNull(dgPendidikanFormal.Columns("Kelulusan").Value) Then cbGradeKelulusan.Text = "" Else cbGradeKelulusan.Text = dgPendidikanFormal.Columns("Kelulusan").Value
    If IsNull(dgPendidikanFormal.Columns("No. Ijazah").Value) Then txtNoIjazah.Text = "" Else txtNoIjazah.Text = dgPendidikanFormal.Columns("No. Ijazah").Value
    If IsNull(dgPendidikanFormal.Columns("Tgl. Ijazah").Value) Then dtpTglIjazah.Value = Null Else dtpTglIjazah.Value = dgPendidikanFormal.Columns("Tgl. Ijazah").Value
    If IsNull(dgPendidikanFormal.Columns("TTD Ijazah").Value) Then txtTandaTanganIjazah.Text = "" Else txtTandaTanganIjazah.Text = dgPendidikanFormal.Columns("TTD Ijazah").Value
    If IsNull(dgPendidikanFormal.Columns("Alamat Sekolah").Value) Then txtAlamatPendidikan.Text = "" Else txtAlamatPendidikan.Text = dgPendidikanFormal.Columns("Alamat Sekolah").Value
    If IsNull(dgPendidikanFormal.Columns("Pimpinan Sekolah").Value) Then txtPimpinanPendidikan.Text = "" Else txtPimpinanPendidikan.Text = dgPendidikanFormal.Columns("Pimpinan Sekolah").Value
    If IsNull(dgPendidikanFormal.Columns("Keterangan").Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgPendidikanFormal.Columns("Keterangan").Value
End Sub

Private Sub dtpTglIjazah_Change()
    dtpTglIjazah.MaxDate = Now
End Sub

Private Sub dtpTglLulus_Change()
    dtpTglLulus.MaxDate = Now
End Sub

Private Sub dtpTglLulus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtIPK.SetFocus
End Sub

Private Sub dtpTglMasuk_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then dtpTglLulus.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call SetComboPendidikan
    Call SetComboGradeKelulusan
    Call subLoadPendidikanFormal
    dtpTglMasuk.Value = Now
    dtpTglLulus.Value = Now
End Sub

Private Sub subLoadPendidikanFormal()
    On Error GoTo hell
    strSQL = "SELECT [No. Urut], Pendidikan, [Nama Sekolah], Jurusan, [Tgl. Masuk], [Tgl. Lulus], IPK, Kelulusan, [No. Ijazah], [Tgl. Ijazah], [TTD Ijazah], " & _
    " [Alamat Sekolah] , [Pimpinan Sekolah], Keterangan, [Nama User] " & _
    " FROM v_RiwayatPendidikanFormal WHERE [ID Peg]='" & mstrIdPegawai & "' ORDER BY [No. Urut]"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgPendidikanFormal
        Set .DataSource = rs
        .Columns("No. Urut").Width = 750
        .Columns("Pendidikan").Width = 2000
        .Columns("Nama Sekolah").Width = 2000
        .Columns("Jurusan").Width = 1000
        .Columns("Tgl. Masuk").Width = 1000
        .Columns("Tgl. Masuk").Caption = "Tgl. Masuk"
        .Columns("Tgl. Lulus").Width = 1000
        .Columns("Tgl. Lulus").Caption = "Tgl. Lulus"
        .Columns("IPK").Width = 500
        .Columns("Kelulusan").Width = 1000
        .Columns("No. Ijazah").Width = 1000
        .Columns("Tgl. Ijazah").Width = 1000
        .Columns("TTD Ijazah").Width = 1000
        .Columns("Alamat Sekolah").Width = 2500
        .Columns("Pimpinan Sekolah").Width = 1500
        .Columns("Keterangan").Width = 2500
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    dcpendidikan.Text = ""
    txtFakultasJurusan.Text = ""
    dtpTglMasuk.Value = Format(Now, "dd/mmmm/yyyy")
    dtpTglLulus.Value = Format(Now, "dd/mmmm/yyyy")
    cbGradeKelulusan.Text = ""
    txtNoIjazah.Text = ""
    dtpTglIjazah.Value = Format(Now, "dd/mmmm/yyyy")
    txtAlamatPendidikan.Text = ""
    txtPimpinanPendidikan.Text = ""
    txtTandaTanganIjazah.Text = ""
    txtNamaPendidikan.Text = ""
    txtIPK.Text = ""
    txtKeterangan = ""
End Sub

Sub SetComboPendidikan()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "Select * from Pendidikan where statusenabled='1'", dbConn, , adLockOptimistic
    Set dcpendidikan.RowSource = rs
    dcpendidikan.ListField = rs.Fields(1).Name
    dcpendidikan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub SetComboGradeKelulusan()
    With cbGradeKelulusan
        .AddItem "Cukup Memuaskan"
        .AddItem "Memuaskan"
        .AddItem "Sangat Memuaskan"
        .AddItem "Cum Laude"
        .ListIndex = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatPendidikanFormal
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtFakultasJurusan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglMasuk.SetFocus
End Sub

Private Sub txtIPK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cbGradeKelulusan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then KeyAscii = 0
End Sub

Private Sub cbGradeKelulusan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNoIjazah.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFakultasJurusan.SetFocus
End Sub

Private Sub txtNoIjazah_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dtpTglIjazah.SetFocus
End Sub

Private Sub dtpTglIjazah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTandaTanganIjazah.SetFocus
End Sub

Private Sub txtAlamatPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtPimpinanPendidikan.SetFocus
End Sub

Private Sub txtPimpinanPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtTandaTanganIjazah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAlamatPendidikan.SetFocus
End Sub
