VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMutasiPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Mutasi Kepegawaian"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "frmMutasiPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
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
      TabIndex        =   2
      Top             =   7200
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
      TabIndex        =   5
      Top             =   7200
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
      TabIndex        =   3
      Top             =   7200
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
      TabIndex        =   4
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
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
      TabIndex        =   6
      Top             =   1080
      Width           =   10455
      Begin VB.TextBox txtnoUrut 
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
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtJabatan 
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
         TabIndex        =   10
         Top             =   1320
         Width           =   4455
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
         Left            =   240
         TabIndex        =   0
         Top             =   2040
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
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
         Format          =   120717315
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Urut"
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
         TabIndex        =   14
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan"
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
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
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
         TabIndex        =   9
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tempat"
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
         TabIndex        =   7
         Top             =   1800
         Width           =   540
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
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
   Begin MSDataGridLib.DataGrid dgRiwayatSIP 
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5953
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8880
      Picture         =   "frmMutasiPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMutasiPegawai.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMutasiPegawai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMutasiPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    txtJabatan.Text = ""
    txtTempat.Text = ""
End Sub

Private Sub cmdHapus_Click()
    If txtNoUrut.Text = "" Then Exit Sub
    
    strSQL = "delete from RiwayatMutasiPegawai  where IdPegawai ='" & mstrIdPegawai & "' and  NoUrut='" & txtNoUrut.Text & "'"
    Call msubRecFO(rs, strSQL)
    
    MsgBox "Hapus berhasil ", vbInformation, "..:."
    Call subLoadData
End Sub

Private Sub cmdSimpan_Click()
    
    'SELECT     nourut,IdPegawai, Jabatan, Tempat, Tahun FROM         RiwayatMutasiPegawai
    
    If txtJabatan.Text = "" Then MsgBox "Nama Jabatan belum di isi", vbInformation, "..:."
    If txtTempat.Text = "" Then MsgBox "Nama Tempat belum di isi", vbInformation, "..:."
    
    strSQL = "SELECT * From RiwayatMutasiPegawai where IdPegawai='" & mstrIdPegawai & "' and NoUrut='" & txtNoUrut.Text & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount = 0 Then
        txtNoUrut.Text = 0
        strSQL = "select max(cast(NoUrut as integer)) from RiwayatMutasiPegawai Where IdPegawai = '" & mstrIdPegawai & "'  "
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then txtNoUrut.Text = IIf(IsNull(rs(0)), "0", rs(0))
        txtNoUrut.Text = Format(Val(txtNoUrut.Text) + 1, "0#")
    
        strSQL = "insert into RiwayatMutasiPegawai values ('" & mstrIdPegawai & "','" & txtNoUrut.Text & "','" & txtJabatan.Text & "'," & _
                 "'" & txtTempat.Text & "','" & Format(dtpTglAkhir.Value, "yyyy") & "')"
        Call msubRecFO(rs, strSQL)
    Else
        strSQL = "update RiwayatMutasiPegawai set Jabatan='" & txtJabatan.Text & "'," & _
                 "Tempat='" & txtTempat.Text & "',Tahun='" & Format(dtpTglAkhir.Value, "yyyy") & "' where IdPegawai ='" & mstrIdPegawai & "' and  NoUrut='" & txtNoUrut.Text & "'"
        Call msubRecFO(rs, strSQL)
    End If
    
    Call subLoadData
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgRiwayatSIP_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    'dtpSK.Value = dgRiwayatSIP.Columns("TglSK")
    dtpTglAkhir.Value = dgRiwayatSIP.Columns("Tahun")
    txtJabatan.Text = dgRiwayatSIP.Columns("Jabatan")
    txtTempat.Text = dgRiwayatSIP.Columns("Tempat")
    txtNoUrut.Text = dgRiwayatSIP.Columns("NoUrut")
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadData
End Sub


Private Sub subClearData()
    dtpTglAkhir.Value = Date
    txtJabatan.Text = ""
    txtTempat.Text = ""
    
End Sub

Private Sub subLoadData()
    strSQL = "SELECT * FROM RiwayatMutasiPegawai where idpegawai='" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    Set dgRiwayatSIP.DataSource = rs
End Sub
