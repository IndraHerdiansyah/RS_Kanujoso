VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSuratSIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Surat Izin Praktek"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "frmSuratSIP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
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
      TabIndex        =   4
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
      TabIndex        =   7
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
      TabIndex        =   5
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
      TabIndex        =   6
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
      TabIndex        =   8
      Top             =   1080
      Width           =   10455
      Begin VB.TextBox txtTTD 
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
         Left            =   4800
         TabIndex        =   19
         Top             =   1320
         Width           =   5415
      End
      Begin VB.TextBox txtNoSTR 
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
         TabIndex        =   17
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtNmPraktik 
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
         TabIndex        =   1
         Top             =   600
         Width           =   5055
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
         TabIndex        =   3
         Top             =   1920
         Width           =   9975
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   600
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
         CustomFormat    =   "dd MMM yyyy HH:mm"
         Format          =   111607811
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin MSComCtl2.DTPicker dtpSK 
         Height          =   330
         Left            =   3120
         TabIndex        =   15
         Top             =   600
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
         CustomFormat    =   "dd MMM yyyy HH:mm"
         Format          =   111607811
         UpDown          =   -1  'True
         CurrentDate     =   38448
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
         Left            =   4800
         TabIndex        =   20
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nomor STR"
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
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Akhir Berlaku"
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
         Left            =   3120
         TabIndex        =   16
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label3 
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
         Left            =   1080
         TabIndex        =   14
         Top             =   360
         Width           =   435
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
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Untuk Praktik"
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
         TabIndex        =   10
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Lokasi"
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
         Top             =   1680
         Width           =   435
      End
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
      Picture         =   "frmSuratSIP.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmSuratSIP.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmSuratSIP.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmSuratSIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHapus_Click()
    If txtNoUrut.Text = "" Then Exit Sub
    
    strSQL = "delete RiwayatIzinPraktek  where IdPegawai ='" & mstrIdPegawai & "' and  NoUrut='" & txtNoUrut.Text & "'"
    Call msubRecFO(rs, strSQL)
    
    Call subLoadData
    MsgBox "Hapus berhasil ", vbInformation, "..:."
End Sub

Private Sub cmdSimpan_Click()
    'SELECT IdPegawai, NoUrut, NmPraktik, NoSK, TglSK, TglBerlaku, TandaTanganSK, Lokasi, IdUser From RiwayatIzinPraktek
    
    If txtNmPraktik.Text = "" Then MsgBox "Nama Praktik belum di isi", vbInformation, "..:."
    If txtNoSTR.Text = "" Then MsgBox "No STR belum di isi", vbInformation, "..:."
    
    strSQL = "SELECT * From RiwayatIzinPraktek where IdPegawai='" & mstrIdPegawai & "' and NoUrut='" & txtNoUrut.Text & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount = 0 Then
        txtNoUrut.Text = 0
        strSQL = "select max(cast(NoUrut as integer)) from RiwayatPangkat Where IdPegawai = '" & mstrIdPegawai & "'  "
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then txtNoUrut.Text = IIf(IsNull(rs(0)), "0", rs(0))
        txtNoUrut.Text = Format(Val(txtNoUrut.Text) + 1, "0#")
    
        strSQL = "insert into RiwayatIzinPraktek values ('" & mstrIdPegawai & "','" & txtNoUrut.Text & "','" & txtNmPraktik.Text & "'," & _
                 "'" & txtNoSTR.Text & "','" & Format(dtpSK.Value, "yyyy-MM-dd 00:00") & "','" & Format(dtpTglAkhir.Value, "yyyy-MM-dd 00:00") & "','" & txtTTD.Text & "'," & _
                 "'" & txtAlamat.Text & "','" & strIDPegawai & "')"
        Call msubRecFO(rs, strSQL)
    Else
        strSQL = "update RiwayatIzinPraktek set NmPraktik='" & txtNmPraktik.Text & "'," & _
                 "NoSK='" & txtNoSTR.Text & "',TglSK='" & Format(dtpSK.Value, "yyyy-MM-dd 00:00") & "',TglBerlaku='" & Format(dtpTglAkhir.Value, "yyyy-MM-dd 00:00") & "',TandaTanganSK='" & txtTTD.Text & "'," & _
                 "Lokasi='" & txtAlamat.Text & "',IdUser='" & strIDPegawai & "' where IdPegawai ='" & mstrIdPegawai & "' and  NoUrut='" & txtNoUrut.Text & "'"
        Call msubRecFO(rs, strSQL)
    End If
    
    Call subLoadData
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgRiwayatSIP_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    dtpSK.Value = dgRiwayatSIP.Columns("TglSK")
    dtpTglAkhir.Value = dgRiwayatSIP.Columns("TglBerlaku")
    txtNoUrut.Text = dgRiwayatSIP.Columns("noUrut")
    txtNmPraktik.Text = dgRiwayatSIP.Columns("NmPraktik")
    txtNoSTR.Text = dgRiwayatSIP.Columns("NoSK")
    txtTTD.Text = dgRiwayatSIP.Columns("TandaTanganSK")
    txtAlamat.Text = dgRiwayatSIP.Columns("Lokasi")
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadData
End Sub


Private Sub subClearData()
    dtpTglAkhir.Value = Date
    dtpSK.Value = Date
    txtNmPraktik.Text = ""
    txtNoSTR.Text = ""
    txtTTD.Text = ""
    txtAlamat.Text = ""
    
End Sub

Private Sub subLoadData()
    strSQL = "SELECT IdPegawai, NoUrut, NmPraktik, NoSK, TglSK, TglBerlaku, TandaTanganSK, Lokasi, IdUser From RiwayatIzinPraktek where idpegawai='" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    Set dgRiwayatSIP.DataSource = rs
End Sub
